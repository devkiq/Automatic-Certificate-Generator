import os
import smtplib
import openpyxl
from email.message import EmailMessage
from PIL import Image, ImageDraw, ImageFont
from dotenv import load_dotenv
import re
import unicodedata

# 1) Carrega variáveis de ambiente do .env
load_dotenv()

# 2) Mapeamento de cada curso ao seu template de certificado
TEMPLATE_MAP = {
    "universidade da vida": "./certificados_curso/certificado_uv.jpg",
    "dons": "./certificados_curso/certificado_dons.jpg",
    'voluntarios': "./certificados_curso/certificado_volunt.jpg",
    'noivos': "./certificados_curso/certificado_noivos.jpg",
    'lideres': "./certificados_curso/certificado_lideres.jpg",
    'intercessao': "./certificados_curso/certificado_intercessão.jpg",
    'inteligencia emocional': "./certificados_curso/certificado_intemocional.jpg",
    'escatologia': "./certificados_curso/certificado_escatologia.jpg",
    'comunicacao na perspectiva biblica': "./certificados_curso/certificado_comoersbibl.jpg",
    'capelania': "./certificados_curso/certificado_capelania.jpg",
    'bem casados': "./certificados_curso/certificado_bemcasados.jpg",
    'ativacao profetica': "./certificados_curso/certificado_ativprofeti.jpg",
    'educacao financeira': "./certificados_curso/certificado_educfinanceira.jpg",
    # adicione aqui outros cursos: "nome_chave": "caminho/para/template.jpg"
}

# 3) Constantes de configuração
SHEET_PATH   = "./planilha_alunos/alunos.xlsx"
FONT_PATH    = "./fonts/SHOWG.TTF"
FONT_SIZE    = 35
OUTPUT_DIR   = "./certificados_gerados"
TEXT_Y       = 330

# 4) Configuração do servidor SMTP de teste local
SMTP_HOST = "smtp.gmail.com"         # ou outro provedor como smtp.office365.com
SMTP_PORT = 465                      # 587 para STARTTLS ou 465 para SSL
SENDER_EMAIL = os.getenv("EMAIL_REMETENTE")
SENDER_PASS  = os.getenv("SENHA_REMETENTE")

# ---------------------- FUNÇÕES ----------------------

def ensure_dir(path):
    os.makedirs(path, exist_ok=True)

def load_sheet(path):
    try:
        wb = openpyxl.load_workbook(path)
        return wb.active
    except FileNotFoundError:
        print(f"Erro: planilha '{path}' não encontrada.")
        exit()

def load_font(path, size):
    try:
        return ImageFont.truetype(path, size)
    except IOError:
        print(f"Erro: fonte '{path}' não encontrada.")
        exit()

def is_valid_email(email):
    return bool(re.match(r"[^@]+@[^@]+\.[^@]+", email))

def center_text(draw, text, font, img_width):
    bbox = draw.textbbox((0,0), text, font=font)
    text_w = bbox[2] - bbox[0]
    x = (img_width - text_w) // 2
    return x, TEXT_Y

def create_certificate(name, course, font):
    key = course.lower().strip()
    template = TEMPLATE_MAP.get(key)
    if not template:
        print(f"⚠️ Sem template para curso '{course}'. Pulando.")
        return None

    img = Image.open(template)
    draw = ImageDraw.Draw(img)
    w, _ = img.size

    text = f"{name}"
    x, y = center_text(draw, text, font, w)
    draw.text((x, y), text, fill="black", font=font)

    filename = f"{OUTPUT_DIR}/{name.replace(' ','_')}_{key}.png"
    img.save(filename)
    return filename

def send_email(recipient, name, attachments):
    msg = EmailMessage()
    msg["Subject"] = "Seus Certificados"
    msg["From"] = SENDER_EMAIL
    msg["To"] = recipient
    msg.set_content(f"Olá {name},\n\nSegue em anexo o(s) seu(s) certificado(s).\n\nDeus te abençoe!")

    for filepath in attachments:
        with open(filepath, "rb") as f:
            file_data = f.read()
            filename = os.path.basename(filepath)
            msg.add_attachment(file_data, maintype="image", subtype="png", filename=filename)

    try:
        with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT) as smtp:
            smtp.login(SENDER_EMAIL, SENDER_PASS)
            smtp.send_message(msg)
            print(f"✅ E‑mail enviado para {recipient}")
    except Exception as e:
        print(f"❌ Falha ao enviar e‑mail para {recipient}: {e}")

def main():
    ensure_dir(OUTPUT_DIR)
    sheet = load_sheet(SHEET_PATH)
    font  = load_font(FONT_PATH, FONT_SIZE)

    for row in sheet.iter_rows(min_row=2, values_only=True):
        nome  = row[1]  # Coluna B
        curso = row[2]  # Coluna C
        email = row[3]  # Coluna D

        if not (nome and email and curso):
            print("⚠️ Linha incompleta. Pulando.")
            continue
        if not is_valid_email(email):
            print(f"⚠️ E‑mail inválido: {email}. Pulando.")
            continue

        # Separa cursos por vírgula, barra ou quebra de linha
        cursos = re.split(r'[,/\n]+', curso)
        cursos = [c.strip() for c in cursos if c.strip()]
        print(f"📚 {nome} está matriculado em: {', '.join(cursos)}")

        arquivos = []
        for c in cursos:
            cert = create_certificate(nome, c, font)
            if cert:
                arquivos.append(cert)

        if arquivos:
            send_email(email, nome, arquivos)

if __name__ == "__main__":
    main()
