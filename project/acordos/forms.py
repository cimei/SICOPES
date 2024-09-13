"""

.. topic:: Acordos (formulários)

   O formulário do módulo *Acordos* recebe dados informados pelo usuário para o registro
   de um novo acordo e é o mesmo utilizado quando da atualização de dados de um acordo já existente.

   * AcordoForm: registrar ou atualizar dados de um acordo.
   * Programa_CNPqForm: registrar ou atualizar dados de um programa do CNPq.
   * ArquivoForm: permite escolher o arquivo excel para carga de dados de acordo.
   * ListaForm: escolher coordenação

**Campos definidos no formulário (todos são obrigatórios):**

"""

# forms.py dentro de acordos

from flask_wtf import FlaskForm
from wtforms import StringField, IntegerField, DecimalField, SubmitField, SelectField, SelectMultipleField
from wtforms.fields import DateField
from wtforms.validators import DataRequired, Regexp, Optional
from flask_wtf.file import FileField, FileAllowed, FileRequired
from project import db
from project.models import Programa_CNPq, Processo_Mae, Coords

class ProgAcordoForm(FlaskForm):

    programa_cnpq = SelectMultipleField('Programa CNPq:')

    submit        = SubmitField('Associar')

class ChamadaAcordoForm(FlaskForm):

    chamada = SelectMultipleField('Chamada:')

    submit  = SubmitField('Associar')    


class AcordoForm(FlaskForm):

    #programa_cnpq    = SelectMultipleField('Programa CNPq:')
    nome             = StringField('Edição/Sigla:',validators=[DataRequired(message="Informe um nome ou edição!")])
    desc             = StringField('Descrição:')
    unid             = SelectField('Unidade:', validators=[DataRequired(message="Escolha uma unidade organizacional!")])
    sei              = StringField('Número SEI:',validators=[DataRequired(message="Informe o Programa!")]) # incluir regex para sei
    epe              = StringField('Sigla da EP:',validators=[DataRequired(message="Informe a Instituição!")])
    uf               = StringField('UF (sigla):',validators=[DataRequired(message="Informe a sigla da UF!")])
    data_inicio      = DateField('Data de início:',format='%Y-%m-%d', validators=(Optional(),))#,validators=[DataRequired(message="Informe data do início!")])
    data_fim         = DateField('Data de término:',format='%Y-%m-%d', validators=(Optional(),))#,validators=[DataRequired(message="Informe data do término!")])
    valor_cnpq       = StringField('Valor CNPq:',validators=[DataRequired(message="Informe o valor!")])
    valor_epe        = StringField('Valor EP:',validators=[DataRequired(message="Informe o valor!")])
    situacao         = SelectField('Situação:')
    capital          = StringField('Capital:',validators=[DataRequired(message="Informe o valor destinado a capital!")])
    custeio          = StringField('Custeio:',validators=[DataRequired(message="Informe o valor destinado a custeio!")])
    bolsas           = StringField('Bolsas:',validators=[DataRequired(message="Informe o valor destinado a bolsas!")]) 
    siafi            = StringField('SIAFI:')                             

    submit           = SubmitField('Registrar')

#
class Programa_CNPqForm(FlaskForm):

    cod_programa   = StringField('Código:',validators=[DataRequired(message="Informe o código do programa!")])
    nome_programa  = StringField('Nome:',validators=[DataRequired(message="Informe o nome do Programa!")])
    sigla_programa = StringField('Sigla:',validators=[DataRequired(message="Informe a sigla do Programa!")])
    coord          = SelectField('Unidade:', validators=[DataRequired(message="Escolha uma Coordenção!")])

    submit      = SubmitField('Registrar')

class ArquivoForm(FlaskForm):

    arquivo = FileField('Arquivo:', validators=[FileRequired(message="Selecione um arquivo!"),FileAllowed(['xls'], 'Somente .xls!')])

    submit  = SubmitField('Importar')

class Altera_proc_mae_Form(FlaskForm):

    coordenador = StringField('Coordenador:',validators=[DataRequired(message="Informe o nome do coordenador ou que não há registro!")])
    situ_mae    = StringField('Situação:',validators=[DataRequired(message="Informe a situação atual!")])

    submit  = SubmitField('Registrar')

class Inclui_proc_mae_Form(FlaskForm):

    proc_mae      = StringField('Processo_Mãe:',validators=[DataRequired(message="Informe número do processo_mãe!")])
    coordenador   = StringField('Coordenador:',validators=[DataRequired(message="Informe o nome do coordenador ou que não há registro!")])
    inic_mae      = DateField('Data de início:',format='%Y-%m-%d', validators=(Optional(),))
    term_mae      = DateField('Data de término:',format='%Y-%m-%d', validators=(Optional(),))
    situ_mae      = StringField('Situação:')

    submit  = SubmitField('Registrar')    

#
def func_ProcMae_Acordo(programas):

    class ProcMae_Acordo(FlaskForm):
        pass

    procs_mae = db.session.query(Processo_Mae.id,Processo_Mae.proc_mae)\
                          .filter(Processo_Mae.cod_programa.in_(programas))\
                          .all()

    lista_procs = [(str(proc[0]),proc[1]) for proc in procs_mae]

    proc_mae  = SelectMultipleField('Processos Mãe:',choices= lista_procs)

    submit      = SubmitField('Registrar')
    setattr(ProcMae_Acordo, "proc_mae", proc_mae)
    setattr(ProcMae_Acordo, "submit", submit)

    return ProcMae_Acordo()

#
class EscolheMaeForm(FlaskForm):

    mae    = SelectMultipleField('Processo:')
    submit = SubmitField('Associar')

# form para escolher a coordenação na lista de acordos
class ListaForm(FlaskForm):

    coords = db.session.query(Coords.sigla)\
                      .order_by(Coords.sigla).all()
    lista_coords = [(c[0],c[0]) for c in coords]
    lista_coords.insert(0,('',''))

    coord        = SelectField('Coordenação:',choices= lista_coords)
    submit       = SubmitField('Filtrar coordenação')

#
class HomologadoForm(FlaskForm):

    prioridade = StringField('Prioridade/Posição:')
    nota       = StringField('Nota:')
    cpf        = StringField('CPF:')
    nome       = StringField('Nome:')
    mod        = StringField('Modalidade:')
    niv        = StringField('Nível:')
    titulo     = StringField('Título:')
    area       = StringField('Área:')
    valor      = StringField('Valor:')

    submit     = SubmitField('Registrar')
