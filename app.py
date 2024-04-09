from flask import Flask, render_template, redirect, url_for, request, send_file, jsonify
from flask_mail import Mail, Message
from flask_sqlalchemy import SQLAlchemy
from sqlalchemy import func
from datetime import datetime
import io
import pyexcel as p
import os

app = Flask(__name__)
uri = os.getenv("DATABASE_URL")
if uri and uri.startswith("postgres://"):
    uri = uri.replace("postgres://", "postgresql://", 1)
app.config["SQLALCHEMY_DATABASE_URI"] = uri
db = SQLAlchemy(app)

# Configurações do Flask-Mail
app.config["MAIL_SERVER"] = "smtp.hostinger.com"
app.config["MAIL_PORT"] = 465
app.config["MAIL_USE_TLS"] = False
app.config["MAIL_USE_SSL"] = True
app.config["MAIL_USERNAME"] = os.getenv("MAIL_USERNAME")
app.config["MAIL_PASSWORD"] = os.getenv("MAIL_PASSWORD")
app.config["MAIL_DEFAULT_SENDER"] = os.getenv("MAIL_DEFAULT_SENDER")
mail = Mail(app)


class Inscricao(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    nome = db.Column(db.String(120), nullable=False)
    cpf = db.Column(db.String(14), nullable=False)
    email = db.Column(db.String(255), nullable=False)
    estado = db.Column(db.String(120), nullable=False)
    cidade = db.Column(db.String(120), nullable=False)
    empresa_instituicao = db.Column(db.String(255), nullable=False)
    cargo = db.Column(db.String(120), nullable=False)
    aceite_termos = db.Column(db.Boolean, default=False, nullable=False)
    data_hora = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)

    def to_dict(self):
        return {
            "id": self.id,
            "nome": self.nome,
            "cpf": self.cpf,
            "email": self.email,
            "estado": self.estado,
            "cidade": self.cidade,
            "empresa_instituicao": self.empresa_instituicao,
            "cargo": self.cargo,
            "aceite_termos": self.aceite_termos,
            "data_hora": self.data_hora.strftime("%Y-%m-%d %H:%M:%S"),
        }


def create_tables():
    with app.app_context():
        db.create_all()


@app.route("/", methods=["GET"])
def index():
    return render_template("formulario.html")


def send_email(to_email, subject, html_content):
    with app.app_context():
        msg = Message(subject, recipients=[to_email], html=html_content)
        mail.send(msg)


@app.route("/inscricao", methods=["POST"])
def add_inscricao():
    nome = request.form["nome"]
    cpf = request.form["cpf"]
    email = request.form["email"]
    estado = request.form["estado"]
    cidade = request.form["cidade"]
    empresa_instituicao = request.form["empresa_instituicao"]
    cargo = request.form["cargo"]
    aceite_termos = request.form.get("aceite_termos") == "on"

    nova_inscricao = Inscricao(
        nome=nome,
        cpf=cpf,
        email=email,
        estado=estado,
        cidade=cidade,
        empresa_instituicao=empresa_instituicao,
        cargo=cargo,
        aceite_termos=aceite_termos,
        data_hora=datetime.utcnow(),
    )

    db.session.add(nova_inscricao)
    db.session.commit()

    # Preparando o conteúdo do e-mail utilizando um template HTML
    to_email = email  # Usa o e-mail fornecido pelo usuário
    subject = "Confirmação"

    # Renderiza o template HTML como string, passando a variável 'responsavel' como 'nome_produtor'
    html_content = render_template("email_template.html", nome=nome)

    send_email(to_email, subject, html_content)

    return redirect(url_for("index", success=True))


@app.route("/inscricoes", methods=["GET"])
def listar_inscricoes():
    inscricoes = Inscricao.query.all()
    return render_template("listar_inscricoes.html", inscricoes=inscricoes)


# Rota para baixar os dados em Excel
@app.route("/download_excel", methods=["GET"])
def download_excel():
    query_sets = Inscricao.query.all()
    data = [inscricao.to_dict() for inscricao in query_sets]

    output = io.BytesIO()
    sheet = p.get_sheet(records=data)
    sheet.save_to_memory("xlsx", output)
    output.seek(0)

    return send_file(
        output,
        as_attachment=True,
        download_name="Inscricoes.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    app.run(debug=True)
