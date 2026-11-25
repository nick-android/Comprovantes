import re
import hashlib
from datetime import datetime, date
from zoneinfo import ZoneInfo

import streamlit as st
import pytesseract
import gspread
from PIL import Image, ImageOps
from google.oauth2.service_account import Credentials

TZ = ZoneInfo("America/Sao_Paulo")  # ajuste se quiser outro fuso
PESSOAS = ["Pessoa 1", "Pessoa 2", "Pessoa 3", "Pessoa 4"]  # coloque os nomes
BANCOS = ["Itaú", "Bradesco", "Santander", "Nubank", "Inter", "Banco do Brasil",
          "Caixa", "Sudameris", "Continental", "Ueno", "Vision", "Familiar",
          "GNB", "Basa", "Atlas", "Mercado Pago", "Tigo Money", "Personal"]
TIPOS = ["Pessoal", "Obra", "Peguei emprestado", "Emprestei"]
COLUNAS = ["Data", "Hora", "Valor", "Moeda", "Banco/Destinatário", "Quem enviou",
           "Tipo", "Obra", "Motivo", "Nº comprovante", "Estado", "Mês"]

st.set_page_config(page_title="Comprovantes", page_icon="🧾")


# ---------- acesso simples ----------
def checar_senha():
    senha = st.secrets.get("app_password")
    if not senha:
        return
    if st.session_state.get("ok"):
        return
    digitada = st.text_input("Senha", type="password")
    if digitada == senha:
        st.session_state["ok"] = True
        st.rerun()
    st.stop()


# ---------- Google Sheets ----------
@st.cache_resource
def abrir_planilha():
    creds = Credentials.from_service_account_info(
        dict(st.secrets["gcp_service_account"]),
        scopes=["https://www.googleapis.com/auth/spreadsheets"],
    )
    return gspread.authorize(creds).open_by_key(st.secrets["sheet_id"])


def preparar_planilha(sh):
    """Cria as abas Registros e Resumo com cores e fórmulas (rodar 1 vez)."""
    try:
        ws = sh.worksheet("Registros")
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet("Registros", rows=2000, cols=len(COLUNAS))
    ws.update("A1", [COLUNAS])
    ws.format("A1:L1", {
        "backgroundColor": {"red": 0.12, "green": 0.25, "blue": 0.45},
        "textFormat": {"bold": True, "foregroundColor": {"red": 1, "green": 1, "blue": 1}},
        "horizontalAlignment": "CENTER",
    })
    ws.freeze(rows=1)
    ws.format("C2:C", {"numberFormat": {"type": "NUMBER", "pattern": "#,##0.00"}})

    cores = {"Pessoal": (0.85, 0.92, 0.83), "Obra": (1, 0.95, 0.8),
             "Peguei emprestado": (0.96, 0.8, 0.8), "Emprestei": (0.81, 0.89, 0.95)}
    regras = []
    for texto, (r, g, b) in cores.items():
        regras.append({"addConditionalFormatRule": {"index": 0, "rule": {
            "ranges": [{"sheetId": ws.id, "startRowIndex": 1, "startColumnIndex": 0,
                        "endColumnIndex": len(COLUNAS)}],
            "booleanRule": {
                "condition": {"type": "CUSTOM_FORMULA",
                              "values": [{"userEnteredValue": f'=$G2="{texto}"'}]},
                "format": {"backgroundColor": {"red": r, "green": g, "blue": b}}}}}})
    sh.batch_update({"requests": regras})

    try:
        rs = sh.worksheet("Resumo")
    except gspread.WorksheetNotFound:
        rs = sh.add_worksheet("Resumo", rows=200, cols=20)
    rs.clear()

    def q(col):
        return (f'=QUERY(Registros!A:L,"select {col}, D, sum(C), count(C) '
                f"where {col}<>'' group by {col}, D order by {col} "
                f"label sum(C) 'Total', count(C) 'Transações', D 'Moeda'\",1)")

    rs.update("A1", [["RESUMO DE COMPROVANTES"]])
    rs.update("A3", [["TOTAL GERAL", "Moeda", "Total", "Transações"],
                     ["Todos", "PYG", '=SUMIFS(Registros!C:C,Registros!D:D,"PYG")',
                      '=COUNTIFS(Registros!D:D,"PYG")'],
                     ["Todos", "BRL", '=SUMIFS(Registros!C:C,Registros!D:D,"BRL")',
                      '=COUNTIFS(Registros!D:D,"BRL")']],
              value_input_option="USER_ENTERED")
    rs.update("A8", [["POR TIPO"]]);   rs.update("A9", [[q("G")]], value_input_option="USER_ENTERED")
    rs.update("F8", [["POR OBRA"]]);   rs.update("F9", [[q("H")]], value_input_option="USER_ENTERED")
    rs.update("K8", [["POR PESSOA"]]); rs.update("K9", [[q("F")]], value_input_option="USER_ENTERED")
    rs.update("P8", [["POR MÊS"]]);    rs.update("P9", [[q("L")]], value_input_option="USER_ENTERED")

    rs.format("A1", {"textFormat": {"bold": True, "fontSize": 16}})
    for rng in ["A3:D3", "A8", "F8", "K8", "P8"]:
        rs.format(rng, {"backgroundColor": {"red": 0.12, "green": 0.25, "blue": 0.45},
                        "textFormat": {"bold": True,
                                       "foregroundColor": {"red": 1, "green": 1, "blue": 1}}})
    rs.format("C4:C5", {"numberFormat": {"type": "NUMBER", "pattern": "#,##0.00"}})


# ---------- leitura do comprovante ----------
def ler_texto(img):
    img = ImageOps.exif_transpose(img).convert("L")
    if img.width < 1200:
        f = 1200 / img.width
        img = img.resize((int(img.width * f), int(img.height * f)))
    img = ImageOps.autocontrast(img)
    return pytesseract.image_to_string(img, lang="spa+por", config="--psm 6")


def para_numero(txt, moeda):
    t = txt.strip().strip(".,")
    if moeda == "PYG":
        t = re.sub(r"[.,]\d{1,2}$", "", t) if re.search(r"[.,]\d{1,2}$", t) and t.count(".") + t.count(",") == 1 and len(t.split(re.findall(r"[.,]", t)[-1])[-1]) <= 2 else t
        return float(re.sub(r"[.,]", "", t) or 0)
    if "," in t:  # formato BR: 1.234,56
        t = t.replace(".", "").replace(",", ".")
    elif t.count(".") > 1:
        t = t.replace(".", "")
    return float(t)


def analisar(texto):
    r = {"valor": None, "moeda": None, "data": None, "banco": None,
         "destino": None, "comprovante": None, "estado": None}
    r["valores_achados"] = []

    padrao = re.compile(r"(Gs\.?|₲|PYG|R\$|BRL)\s*([\d][\d\.\,]*)", re.I)
    for m in padrao.finditer(texto):
        moeda = "BRL" if m.group(1).upper() in ("R$", "BRL") else "PYG"
        try:
            v = para_numero(m.group(2), moeda)
            if v > 0:
                r["valores_achados"].append((v, moeda))
        except ValueError:
            pass
    unicos = list(dict.fromkeys(r["valores_achados"]))
    r["valores_unicos"] = unicos
    if unicos:
        r["valor"], r["moeda"] = unicos[0]

    m = re.search(r"(\d{2})[/\-.](\d{2})[/\-.](\d{4})", texto)
    if m:
        try:
            r["data"] = date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
        except ValueError:
            pass

    for b in BANCOS:
        if re.search(re.escape(b), texto, re.I):
            r["banco"] = b
            break

    m = re.search(r"(?:beneficiario|destinatario|destino|a favor de)\s*[:\-]?\s*([^\n]{3,60})",
                  texto, re.I)
    if m:
        r["destino"] = m.group(1).strip()

    m = re.search(r"(?:comprobante|operaci[oó]n|transacci[oó]n|referencia|n[°ºo]\.?)\s*[:#]?\s*(\d{5,})",
                  texto, re.I)
    if m:
        r["comprovante"] = m.group(1)

    low = texto.lower()
    if "en proceso" in low or "pendiente" in low:
        r["estado"] = "Pendente"
    elif re.search(r"exitos|realizad|aprobad|confirmad|completad", low):
        r["estado"] = "Confirmado"
    return r


def tabela_leitura(r):
    def linha(campo, valor, ok_extra=None, dica_falta=""):
        if valor in (None, ""):
            return {"Campo": campo, "Lido": "—", "Status": "❌ Não lido", "O que fazer": dica_falta}
        if ok_extra:
            return {"Campo": campo, "Lido": str(valor), "Status": "⚠️ Confuso", "O que fazer": ok_extra}
        return {"Campo": campo, "Lido": str(valor), "Status": "✅ Lido", "O que fazer": "Só conferir"}

    multi = len(r["valores_unicos"]) > 1
    return [
        linha("Valor", f'{r["valor"]:,.0f}' if r["valor"] else None,
              "Vi mais de um valor, confira qual é o certo" if multi else None,
              "Digite o valor abaixo"),
        linha("Moeda", r["moeda"], None, "Escolha PYG ou BRL"),
        linha("Data", r["data"].strftime("%d/%m/%Y") if r["data"] else None, None,
              "Escolha a data abaixo"),
        linha("Banco", r["banco"], None, "Escolha ou escreva o banco/empresa"),
        linha("Destinatário", r["destino"], "Confira se o nome saiu certo", "Escreva para quem foi"),
        linha("Nº comprovante", r["comprovante"], None, "Opcional, ajuda a evitar duplicados"),
        linha("Estado", r["estado"], None, "Confirme se o pagamento foi concluído"),
    ]


# ---------- interface ----------
checar_senha()
st.title("🧾 Enviar comprovante")

arquivo = st.file_uploader("Foto ou print do comprovante", type=["png", "jpg", "jpeg", "webp"])

if arquivo:
    dados_arq = arquivo.getvalue()
    fid = hashlib.md5(dados_arq).hexdigest()[:8]
    if st.session_state.get("fid") != fid:
        with st.spinner("Lendo comprovante..."):
            img = Image.open(arquivo)
            st.session_state["analise"] = analisar(ler_texto(img))
            st.session_state["fid"] = fid
            st.session_state["enviado"] = False
    r = st.session_state["analise"]

    st.image(dados_arq, width=260)
    st.subheader("O que o sistema entendeu")
    st.table(tabela_leitura(r))
    if not r["valor"] and not r["banco"]:
        st.warning("Não consegui ler quase nada. Preencha todos os campos abaixo.")

    st.subheader("Confirme e complete")
    c1, c2 = st.columns([2, 1])
    valor = c1.number_input("Valor", min_value=0.0, step=1.0, format="%.2f",
                            value=float(r["valor"] or 0), key=f"v{fid}")
    moeda = c2.selectbox("Moeda", ["PYG", "BRL"],
                         index=0 if (r["moeda"] or "PYG") == "PYG" else 1, key=f"m{fid}")
    data_comp = st.date_input("Data do comprovante", value=r["data"] or datetime.now(TZ).date(),
                              key=f"d{fid}", format="DD/MM/YYYY")
    banco = st.text_input("Banco/Empresa/Destinatário",
                          value=r["destino"] or r["banco"] or "", key=f"b{fid}")
    estado = st.selectbox("Estado", ["Confirmado", "Pendente"],
                          index=1 if r["estado"] == "Pendente" else 0, key=f"e{fid}")
    comprovante = st.text_input("Nº do comprovante (opcional)",
                                value=r["comprovante"] or "", key=f"n{fid}")

    st.divider()
    quem = st.selectbox("Quem está enviando?", PESSOAS, key=f"q{fid}")
    tipo = st.radio("Finalidade", TIPOS, horizontal=True, key=f"t{fid}")
    obra = ""
    if tipo == "Obra":
        obra = st.text_input("Qual obra?", key=f"o{fid}")
    motivo = st.text_input("Motivo / descrição", key=f"mo{fid}",
                           placeholder="Ex.: cimento, diária do pedreiro, almoço...")

    if st.button("📤 Enviar", type="primary", disabled=st.session_state.get("enviado", False)):
        erros = []
        if valor <= 0: erros.append("informe o valor")
        if not banco.strip(): erros.append("informe o banco/destinatário")
        if tipo == "Obra" and not obra.strip(): erros.append("escreva qual obra")
        if not motivo.strip(): erros.append("escreva o motivo")
        if erros:
            st.error("Falta: " + ", ".join(erros) + ".")
        else:
            try:
                sh = abrir_planilha()
                ws = sh.worksheet("Registros")
                if comprovante.strip():
                    existentes = ws.col_values(10)
                    if comprovante.strip() in existentes:
                        st.error("Esse nº de comprovante já foi registrado.")
                        st.stop()
                agora = datetime.now(TZ)
                ws.append_row([
                    data_comp.strftime("%Y-%m-%d"), agora.strftime("%H:%M"), valor, moeda,
                    banco.strip(), quem, tipo, obra.strip(), motivo.strip(),
                    "'" + comprovante.strip() if comprovante.strip() else "",
                    estado, "'" + data_comp.strftime("%Y-%m"),
                ], value_input_option="USER_ENTERED")
                st.session_state["enviado"] = True
                st.success("✅ Registrado na planilha!")
            except Exception as e:
                st.error(f"Não consegui salvar: {e}")

link = st.secrets.get("sheet_url")
if link:
    st.markdown(f"[📊 Abrir planilha completa]({link})")

with st.expander("⚙️ Preparar planilha (só na primeira vez)"):
    if st.button("Criar abas Registros e Resumo"):
        preparar_planilha(abrir_planilha())
        st.success("Planilha preparada.")
        