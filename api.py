import os
from flask import Flask, jsonify, request
from flask_sqlalchemy import SQLAlchemy
import csv
import io

app = Flask(__name__)

# Configuração do Banco de Dados a partir da variável de ambiente do Railway
app.config['SQLALCHEMY_DATABASE_URI'] = os.environ.get('DATABASE_URL')
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False

db = SQLAlchemy(app)

# --- Modelos do Banco de Dados ---
class Cidade(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    cidade = db.Column(db.String(100), unique=True, nullable=False)
    url = db.Column(db.String(255))
    cnpj_selector = db.Column(db.String(100))
    login_selector = db.Column(db.String(100))
    senha_seletor = db.Column(db.String(100))

class Empresa(db.Model):
    id = db.Column(db.String(20), primary_key=True)
    empresa = db.Column(db.String(150), nullable=False)
    cnpj = db.Column(db.String(20))
    login = db.Column(db.String(100))
    senha = db.Column(db.String(100))
    cidade = db.Column(db.String(100))
    regime = db.Column(db.String(50))

# --- Rotas da API para CIDADES ---
@app.route('/cidades', methods=['GET'])
def get_cidades():
    cidades = Cidade.query.order_by(Cidade.cidade).all()
    return jsonify([{'cidade': c.cidade, 'url': c.url, 'cnpj-selector': c.cnpj_selector, 'login-selector': c.login_selector, 'senha_seletor': c.senha_seletor} for c in cidades])

@app.route('/cidades', methods=['POST'])
def add_cidade():
    data = request.json
    nova_cidade = Cidade(cidade=data['cidade'], url=data['url'], cnpj_selector=data['cnpj-selector'], login_selector=data['login-selector'], senha_seletor=data['senha_seletor'])
    db.session.add(nova_cidade)
    db.session.commit()
    return jsonify({'message': 'Cidade adicionada com sucesso'}), 201

# --- Rotas da API para EMPRESAS ---
@app.route('/empresas', methods=['GET'])
def get_empresas():
    empresas = Empresa.query.order_by(Empresa.empresa).all()
    return jsonify([{'id': e.id, 'empresa': e.empresa, 'cnpj': e.cnpj, 'login': e.login, 'senha': e.senha, 'cidade': e.cidade, 'regime': e.regime} for e in empresas])

@app.route('/empresas', methods=['POST'])
def update_empresas():
    empresas_data = request.json
    try:
        # Apaga todas as empresas existentes
        Empresa.query.delete()
        # Adiciona as novas empresas
        for data in empresas_data:
            nova_empresa = Empresa(id=data['id'], empresa=data['empresa'], cnpj=data['cnpj'], login=data['login'], senha=data['senha'], cidade=data['cidade'], regime=data['regime'])
            db.session.add(nova_empresa)
        db.session.commit()
        return jsonify({'message': 'Lista de empresas atualizada com sucesso'}), 200
    except Exception as e:
        db.session.rollback()
        return jsonify({'error': str(e)}), 500

# --- Rota para popular o banco de dados pela primeira vez ---
@app.route('/seed-database', methods=['POST'])
def seed_database():
    # Verifica se um 'token' simples foi enviado para evitar acesso acidental
    if request.headers.get('X-Seed-Token') != 'agrocontat123':
        return jsonify({'error': 'Não autorizado'}), 401

    try:
        # Limpa tabelas existentes
        db.session.query(Empresa).delete()
        db.session.query(Cidade).delete()
        db.session.commit()

        # Popula cidades
        cidades_file = request.files.get('cidades_csv')
        if cidades_file:
            stream = io.StringIO(cidades_file.stream.read().decode("UTF-8"), newline=None)
            reader = csv.DictReader(stream, delimiter=';')
            for row in reader:
                c = Cidade(cidade=row['cidade'], url=row['url'], cnpj_selector=row['cnpj-selector'], login_selector=row['login-selector'], senha_seletor=row['senha_seletor'])
                db.session.add(c)
        
        # Popula empresas
        empresas_file = request.files.get('empresas_csv')
        if empresas_file:
            stream = io.StringIO(empresas_file.stream.read().decode("UTF-8"), newline=None)
            reader = csv.DictReader(stream, delimiter=';')
            for row in reader:
                e = Empresa(id=row['id'], empresa=row['empresa'], cnpj=row['cnpj'], login=row['login'], senha=row['senha'], cidade=row['cidade'], regime=row['regime'])
                db.session.add(e)

        db.session.commit()
        return jsonify({'message': 'Banco de dados populado com sucesso!'})
    except Exception as e:
        db.session.rollback()
        return jsonify({'error': str(e)}), 500

# Rota para inicializar o banco de dados
@app.route('/init-db', methods=['GET'])
def init_db():
    with app.app_context():
        db.create_all()
    return "Banco de dados inicializado!"

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 5000)))
