import os
from flask import Flask, render_template, request, redirect, url_for, flash
from flask_sqlalchemy import SQLAlchemy
from flask_login import LoginManager, UserMixin, login_user, login_required, logout_user, current_user
from werkzeug.security import generate_password_hash, check_password_hash
import pandas as pd
from onedrive_utils import enviar_para_onedrive, ler_arquivo_onedrive


app = Flask(__name__)
@app.route('/')
def home():
    return redirect(url_for('login'))
#app.config['SECRET_KEY'] = 'chave-secreta'  # Alterar para uma chave segura
app.config['SECRET_KEY'] = os.environ.get('SECRET_KEY')
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///usuarios.db'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
db = SQLAlchemy(app)

login_manager = LoginManager(app)
login_manager.login_view = "login"

# 📂 Caminho do OneDrive
# ONEDRIVE_FOLDER = r'C:/Users/ewila/OneDrive - PRODAM Office 365/Dados-comitê'
# os.makedirs(ONEDRIVE_FOLDER, exist_ok=True)
# ONEDRIVE_FOLDER = 'dados'
# os.makedirs(ONEDRIVE_FOLDER, exist_ok=True)

# 📄 Nome do arquivo Excel para armazenar respostas
dados_excel = 'dados_formulario.xlsx'

# 📌 Modelo de usuário
class Usuario(UserMixin, db.Model):
    id = db.Column(db.Integer, primary_key=True)
    nome = db.Column(db.String(100), nullable=False)
    email = db.Column(db.String(100), unique=True, nullable=False)
    senha = db.Column(db.String(200), nullable=False)

@login_manager.user_loader
def load_user(user_id):
    return Usuario.query.get(int(user_id))

# 📌 Rota de Cadastro
@app.route('/cadastro', methods=['GET', 'POST'])
def cadastro():
    if request.method == 'POST':
        nome = request.form['nome']
        email = request.form['email']
        senha = request.form['senha']
        senha_hash = generate_password_hash(senha)

        if Usuario.query.filter_by(email=email).first():
            flash("E-mail já cadastrado!", "danger")
            return redirect(url_for("cadastro"))

        novo_usuario = Usuario(nome=nome, email=email, senha=senha_hash)
        db.session.add(novo_usuario)
        db.session.commit()

        flash("Cadastro realizado com sucesso! Faça login.", "success")
        return redirect(url_for("login"))

    return render_template('cadastro.html')

# 📌 Rota de Login
@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        email = request.form['email']
        senha = request.form['senha']
        usuario = Usuario.query.filter_by(email=email).first()

        if usuario and check_password_hash(usuario.senha, senha):
            login_user(usuario)
            return redirect(url_for("minhas_respostas"))

        flash("Login inválido!", "danger")
        return redirect(url_for("login"))

    return render_template('login.html')


# 📌 Formulário de Respostas (Apenas para usuários logados)
@app.route('/formulario', methods=['GET', 'POST'])
@login_required
def formulario():
    if request.method == 'POST':
        nome_secretaria = request.form.get('nome_secretaria', '')
        nome_responsavel = request.form.get('nome_responsavel', '')
        tipo_acao = request.form.get('tipo_acao', '')
        situacao_problema = request.form.get('situacao_problema', '')
        evento_acao = request.form.get('evento', '')
        niveis_planejamento = request.form.get('nivel_vulnerabilidade1', '')
        municipios_planejamento = request.form.get('municipio_section_planejamento', '')
        objetivos_planejamento = request.form.get('objetivos_planejamento', '')
        valor_empenhado_planejamento = request.form.get('valor_planejamento', '')
        orcamento_planejamento = request.form.get('orcamento_planejamento', '')
        inicio_planejamento = request.form.get('inicio_planejamento', '')
        termino_planejamento = request.form.get('termino_planejamento', '')
        niveis_andamento = request.form.get('nivel_vulnerabilidade2', '')
        municipios_andamento = request.form.get('municipio_section_andamento', '')
        descricao_andamento = request.form.get('descricao_andamento', '')
        valor_empenhado_andamento = request.form.get('valor_empenhado_andamento', '')
        inicio_andamento = request.form.get('inicio_andamento', '')
        termino_andamento = request.form.get('termino_andamento', '')
        desafios_andamento = request.form.get('desafios_andamento', '')
        populacao_andamento = request.form.get('informacao_populacao', '')
        secretarias_andamento = request.form.get('secretarias_andamento', '')
        niveis_realizada = request.form.get('nivel_vulnerabilidade3', '')
        municipios_realizada = request.form.get('municipio_section_realizada', '')
        descricao_realizada = request.form.get('descricao_realizada', '')
        valor_empenhado_realizada = request.form.get('valor_realizada', '')
        inicio_realizada = request.form.get('inicio_realizada', '')
        termino_realizada = request.form.get('termino_realizada', '')
        secretarias_realizada = request.form.get('secretarias_realizada', '')

        # 📄 Adicionar as respostas ao Excel
        novo_dado = pd.DataFrame({
            'Usuário': [current_user.email],
            'Nome Secretaria': [nome_secretaria],
            'Nome Responsavel': [nome_responsavel],
            'Tipo Ação': [tipo_acao],
            'Situação Problema': [situacao_problema],
            'Evento': [evento_acao],
            'Niveis Planejamento': [niveis_planejamento],
            'Municipios Planejamento': [municipios_planejamento],
            'Objetivos Planejamento': [objetivos_planejamento],
            'Empenho Planejamento': [valor_empenhado_planejamento],
            'Orcamento Planejamento': [orcamento_planejamento],
            'Inicio Planejamento': [inicio_planejamento],
            'Termino Planejamento': [termino_planejamento],
            'Niveis Andamento': [niveis_andamento],
            'Municipios Andamento': [municipios_andamento],
            'Descricao Andamento': [descricao_andamento],
            'Empenho Andamento': [valor_empenhado_andamento],
            'Inicio Andamento': [inicio_andamento],
            'Termino Andamento': [termino_andamento],
            'Desafios Andamento': [desafios_andamento],
            'Populacao Andamento': [populacao_andamento],
            'Secretarias Andamento': [secretarias_andamento],
            'Niveis Realizada': [niveis_realizada],
            'Municipios Realizada': [municipios_realizada],
            'Descrição Realizada': [descricao_realizada],
            'Empenho Realizada': [valor_empenhado_realizada],
            'Início Realizada': [inicio_realizada],
            'Término Realizada': [termino_realizada],
            'Secretarias Realizada': [secretarias_realizada]
        })

        try:
            df_final = novo_dado

            # Enviar o arquivo atualizado para o OneDrive
            sucesso = enviar_para_onedrive(df_final)

            if not sucesso:
                flash("Erro ao enviar para o OneDrive", "danger")
                return render_template('formulario.html')  # Evita redirecionamento contínuo

            flash("Dados enviados com sucesso!", "success")
            return redirect(url_for("minhas_respostas"))

        except Exception as e:
            flash(f"Erro inesperado: {e}", "danger")
            return render_template('formulario.html')  # Evita redirecionamento contínuo

    return render_template('formulario.html')  # Para GET, renderiza o formulário

# 📌 Visualizar respostas do usuário logado
@app.route('/minhas_respostas')
@login_required
def minhas_respostas():
    df = ler_arquivo_onedrive(dados_excel)

    if df is not None:
        df_usuario = df[df['Usuário'] == current_user.email]
        respostas = df_usuario.to_dict(orient="records")
    else:
        respostas = []

    return render_template('minhas_respostas.html', respostas=respostas, nome_usuario=current_user.nome)



# 📌 Editar resposta
@app.route('/editar-resposta/<int:indice>', methods=['GET', 'POST'])
@login_required
def editar_resposta(indice):
    df = ler_arquivo_onedrive(dados_excel)

    if df is None:
        flash("Erro ao acessar dados no OneDrive!", "danger")
        return redirect(url_for('minhas_respostas'))

    if 0 <= indice < len(df):
        resposta = df.iloc[indice]

        if resposta["Usuário"] != current_user.email:
            flash("Você não tem permissão para editar esta resposta!", "danger")
            return redirect(url_for('minhas_respostas'))

        if request.method == 'POST':
            df.at[indice, 'Nome Secretaria'] = request.form['nome_secretaria']
            df.at[indice, 'Situação Problema'] = request.form['situacao_problema']

            sucesso = enviar_para_onedrive(df)

            if not sucesso:
                flash("Erro ao salvar dados editados no OneDrive!", "danger")
                return redirect(url_for('minhas_respostas'))

            flash("Resposta editada com sucesso!", "success")
            return redirect(url_for('minhas_respostas'))

        return render_template('editar_resposta.html', resposta=resposta, indice=indice)

    flash("Resposta não encontrada!", "danger")
    return redirect(url_for('minhas_respostas'))

# 📌 Excluir resposta
@app.route('/excluir-resposta/<int:indice>', methods=['POST'])
@login_required
def excluir_resposta(indice):
    df = ler_arquivo_onedrive(dados_excel)

    if df is None:
        flash("Erro ao acessar dados no OneDrive!", "danger")
        return redirect(url_for('minhas_respostas'))

    if 0 <= indice < len(df):
        resposta = df.iloc[indice]

        if resposta["Usuário"] != current_user.email:
            flash("Você não tem permissão para excluir esta resposta!", "danger")
            return redirect(url_for('minhas_respostas'))

        df = df.drop(index=indice).reset_index(drop=True)

        sucesso = enviar_para_onedrive(df)

        if not sucesso:
            flash("Erro ao excluir resposta no OneDrive!", "danger")
            return redirect(url_for('minhas_respostas'))

        flash("Resposta excluída com sucesso!", "success")
        return redirect(url_for('minhas_respostas'))

    flash("Resposta não encontrada!", "danger")
    return redirect(url_for('minhas_respostas'))
        
@app.route('/logout', methods=['POST'])
@login_required  # Garantir que o usuário esteja logado para fazer o logout
def logout():
    logout_user()  # Função do Flask-Login para deslogar o usuário
    flash("Você saiu com sucesso.", "info")  # Mensagem de sucesso ao deslogar
    return redirect(url_for('login'))  # Redireciona para a página de login


# 📌 Rodar o aplicativo
if __name__ == '__main__':
    with app.app_context():
        db.create_all()  # Cria banco de dados se não existir
    app.run(debug=True)
    # app.run(debug=True)

