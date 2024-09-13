"""
.. topic:: Acordos (views)

    Os Acordos são instrumentos de parceria entre o CNPq e Entidades Parceiras Estaduais - EPEs onde
    não há repasse direto de recursos entre as partes.

    O CNPq custeia as bolsas dos contemplados em processos seletivos organizados pelas EPEs
    e estas, a título de contrapartida, custeiam outras despesas dos projetos.

    Toda bolsa é implementada por meio de um processo de bolsa (processo filho), que, por sua vez, deve estar vinculado a
    um processo mãe.

    Em princípio, é no processo mãe que são definidos a quantidade máxima de bolsas que poderão ser implementadas no projeto,
    pois o processo mãe tem um valor de concessão definido, bem como uma vigência que limita as vigências dos processos filho.

    A indicação dos bolsisstas  é feita pela EPE em plataforma específica do CNPq e este módulo trabalha com os
    dados fornecidos por este sistema, via planilha excel enviada pela COSAO, sob demanda da COPES.

    Um acordo tem atributos que são registrados no momento de sua criação. Todos são obrigatórios:

    * Edição do programa ao qual ele está vinculado
    * Número do processo SEI
    * Sigla da EPE
    * Unidade da Federação da EPE
    * Data de início
    * Data de término
    * Valor alocado pelo CNPq
    * Valor alocado pela EPE

    Os valores pagos são calculados pela soma de todos os pagamentos registrados para cada processo filho da planilha COSAO.
    Da mesma forma, é feito o cálculo da quantidade de mensalides pagas.

    Os valores a pagar consistem da multiplicação da quantidade de meses entre a data de referência (data de geração da
    planilha COSAO) e o fim de vigência de cada processo-filho pelo valor da respectiva memsalidade do nível de bolsa no
    qual o bolsista foi enquadrado. Da mesma forma, é feito o cálculo da quantidade de mensalides a pagar.

.. topic:: Ações relacionadas aos acordos

    * Listar acordos por edição do programa: lista_acordos
    * Atualizar dados de um acordo: update
    * Registrar um acordo no sistema: cria_acordo
    * Deletar o registro de um acordo: deleta_acordo
    * Listar demandas de um determinado acordo: acordo_demandas
    * Registrar um programa do CNPq no sistema: cria_programa_cnpq
    * Listar programas do CNPq: lista_programa_cnpq
    * Atualizar programas do CNPq: atualiza_programa_cnpq
    * Lista processos mãe de um acordo: lista_processos_mae_por_acordo
    * Alterada dados de processo_mãe: altera_mae
    * Associar processos mãe a um acordo: processo_mae_acordo
    * Desassociar processo mãe de um acordo: deleta_processo_mae
    * Listar processos filho de um processo mãe: lista_processos_filho
    * Listar bolsistas (cpf) de um processo mãe: lista_bolsistas
    * Listar os processos filho de um acordo: lista_processos_filho_por_acordo
    * Carregar situações de arquivo do sigef: carrega_sit_sigef
    * Listar bolsistas (cpf) de um acordo: lista_bolsistas_acordo
    * Acordos por Programa: resumo_acordos
    * Acordos por UF: brasil_acordos
    * Edições de cada programa: edic_programa
    * Gasto mensal por acordo: gasto_mes
    * Listar processos de uma chamada: processos_chamada

"""

# views.py na pasta acordos

from re import I
from flask import render_template,url_for,flash, redirect,request,Blueprint,send_from_directory
from flask_login import current_user,login_required
from sqlalchemy import func, distinct, not_, or_, cast, String, literal
from sqlalchemy.sql import label
from project import db
from project.models import Acordo, RefCargaPDCTR, PagamentosPDCTR, Processo_Mae, Bolsa, User, Demanda,\
                           Chamadas, Programa_CNPq, Acordo_ProcMae, Processo_Filho, Coords, grupo_programa_cnpq,\
                           capital_custeio,DadosSEI, chamadas_cnpq, chamadas_cnpq_acordos, financeiro_acordo, RefSICONV
from project.acordos.forms import AcordoForm, Programa_CNPqForm, func_ProcMae_Acordo, ListaForm, ArquivoForm,\
                                  Altera_proc_mae_Form, ProgAcordoForm, Inclui_proc_mae_Form, ChamadaAcordoForm,\
                                  EscolheMaeForm
from project.demandas.views import registra_log_auto
from project.core.views import consultaDW, chamadas_DW

import locale
import datetime
from datetime import datetime as dt
from dateutil.rrule import rrule, MONTHLY
import xlrd
import tempfile
from werkzeug.utils import secure_filename
import os
from folium import Map, Circle, Popup
import csv
import re

acordos = Blueprint('acordos',__name__,
                            template_folder='templates/acordos')

#
def none_0(a):
    '''
    DOCSTRING: Transforma None em 0.
    INPUT: campo a ser trandormado.
    OUTPUT: 0, se a entrada for None, caso contrário, a entrada.
    '''
    if a == None:
        a = 0
    return a

def cargaSit(entrada):

    print ('\n')
    print ('<<',dt.now().strftime("%x %X"),'>> ',' Carga de arquivo de situações de processos-filho iniciada...')

    book = xlrd.open_workbook(filename=entrada,ragged_rows=True)
    planilha = book.sheet_by_index(0)

    linha_cabeçalho = planilha.row_values(0, start_colx=0, end_colx=None)

    print ('Planilha: SIGEF')
    print (f'Cabeçalho original: {len(linha_cabeçalho)} campos')
    print (f'Quantidade de registros na planilha: {planilha.nrows - 1 }')
    print ('\n')

    qtd_linhas = planilha.nrows - 1

    for i in range(qtd_linhas):

        linha_base = planilha.row_values(i + 1, start_colx=0, end_colx=None)

        proc = planilha.cell_value(i + 1, linha_cabeçalho.index('Processo'))
        sit  = planilha.cell_value(i + 1, linha_cabeçalho.index('Situação'))

        processo_filho = db.session.query(Processo_Filho).filter(Processo_Filho.processo == proc).all()

        for p in processo_filho:

            if p.situ_filho != sit:
                p.situ_filho = sit

        db.session.commit()

        pag_PDCTR = db.session.query(PagamentosPDCTR).filter(PagamentosPDCTR.processo == proc).all()

        for p in pag_PDCTR:

            if p.situ_filho != sit:
                p.situ_filho = sit

        db.session.commit()

    print ('Carga finalizada!')
    print ('\n')

#
def cria_csv(arq,linha,tabela):
  '''Recebe caminho do arquivo como string, campos da tabela como lista e a tabela propriamente dita'''
  with open(arq,'w',encoding='UTF8',newline='') as f:
        writer = csv.writer(f, delimiter=';')
        writer.writerow(linha)
        writer.writerows(tabela)

@acordos.route('/<lista>/<coord>/lista_acordos', methods=['GET', 'POST'])
def lista_acordos(lista,coord):
    """
    +---------------------------------------------------------------------------------------+
    |Apresenta uma lista dos acordos por edição do programa.                                |
    |                                                                                       |
    |Ao varrer a tabela de acordos, resgatando os cadastrados na edição consultada, são     |
    |feitas consultas aos dados envidados pela COSAO e também aos                           |
    |dados sobre os respectivos processos-mãe relacionados a cada acordo para que se        |
    |calcule a quantidade de bolsistas contemplados (CPFs), a quantidade de bolsas          |
    |implementadas (processos-filho), a quantidade de mensalidades pagas, o valor pago,     |
    |o valor a pagar e o saldo (valor alocado pelo CNPq no acordo - valor pago - valor a    |
    |pagar).                                                                                |
    |                                                                                       |
    |No topo da tela há a opção de se inserir um novo acordo e o número sequencial de cada  |
    |acordo (#), ao ser clicado, permite que seus dados possam ser editados.                |
    |                                                                                       |
    +---------------------------------------------------------------------------------------+
    """

    hoje = dt.today()

    # pega unidade do usuário logado
    unidade = db.session.query(User.coord).filter(User.id==current_user.id).first()
    
    form = ListaForm()

    if form.validate_on_submit():

        coord = form.coord.data

        if coord == '' or coord == None:
            coord = '*'

        return redirect(url_for('acordos.lista_acordos',lista=lista,coord=coord))

    else:

        ## lê data de carga de chamadas do DW
        data_carga = db.session.query(RefSICONV.data_cha_dw).first()
        data_cha = data_carga.data_cha_dw

        if coord == '*':

            form.coord.data = ''
            unid = '%'

        elif coord == 'usu':

            form.coord.data = unidade.coord

            # se unidade for pai, pega filhos
            filhos = db.session.query(Coords.sigla).filter(Coords.pai == unidade.coord).all()
            l_filhos = [f.sigla for f in filhos]
            l_filhos.append(unidade.coord)

            if filhos:
                unid = l_filhos
            else:
                unid = unidade.coord
        
        else:
            
            form.coord.data = coord

            # se unidade for pai, pega filhos
            filhos = db.session.query(Coords.sigla).filter(Coords.pai == coord).all()
            l_filhos = [f.sigla for f in filhos]
            l_filhos.append(coord)

            if filhos:
                unid = l_filhos               
            else:
                unid = coord


        # contabiliza quantidade de programas por acordo
        cont_prog = db.session.query(grupo_programa_cnpq.id_acordo,
                                     label('qtd_prog',func.count(grupo_programa_cnpq.id_programa)))\
                              .group_by(grupo_programa_cnpq.id_acordo)\
                              .subquery()

        # contabiliza quantidade de chamadas por acordo
        cont_cham = db.session.query(chamadas_cnpq_acordos.acordo_id,
                                     label('qtd_cha',func.count(chamadas_cnpq_acordos.id)))\
                               .group_by(chamadas_cnpq_acordos.acordo_id)\
                               .subquery()                      

        if lista == 'todos':
            if type(unid) is str:
                # acordos onde o nome da unidade está no campo unidade_cnpq
                acordos_v = db.session.query(label('id',distinct(Acordo.id)),
                                            Acordo.nome,
                                            Acordo.sei,
                                            Acordo.epe,
                                            Acordo.uf,
                                            Acordo.data_inicio,
                                            Acordo.data_fim,
                                            Acordo.valor_cnpq,
                                            Acordo.valor_epe,
                                            label('unid',Acordo.unidade_cnpq),
                                            Acordo.situ,
                                            Acordo.desc,
                                            cont_prog.c.qtd_prog,
                                            Acordo.siafi,
                                            cont_cham.c.qtd_cha,
                                            Acordo.capital,
                                            Acordo.custeio,
                                            Acordo.bolsas)\
                                    .outerjoin(cont_cham,cont_cham.c.acordo_id == Acordo.id)\
                                    .outerjoin(cont_prog,cont_prog.c.id_acordo == Acordo.id)\
                                    .filter(Acordo.unidade_cnpq.like(unid))\
                                    .order_by(Acordo.situ.desc(),Acordo.data_fim,Acordo.nome,Acordo.epe).all()

            elif type(unid) is list: 
                # acordos onde o nome da unidade está no campo unidade_cnpq
                acordos_v = db.session.query(label('id',distinct(Acordo.id)),
                                             Acordo.nome,
                                             Acordo.sei,
                                             Acordo.epe,
                                             Acordo.uf,
                                             Acordo.data_inicio,
                                             Acordo.data_fim,
                                             Acordo.valor_cnpq,
                                             Acordo.valor_epe,
                                             label('unid',Acordo.unidade_cnpq),
                                             Acordo.situ,
                                             Acordo.desc,
                                             cont_prog.c.qtd_prog,
                                             Acordo.siafi,
                                             cont_cham.c.qtd_cha,
                                            Acordo.capital,
                                            Acordo.custeio,
                                            Acordo.bolsas)\
                                    .outerjoin(cont_cham,cont_cham.c.acordo_id == Acordo.id)\
                                    .outerjoin(cont_prog,cont_prog.c.id_acordo == Acordo.id)\
                                    .filter(Acordo.unidade_cnpq.in_(unid))\
                                    .order_by(Acordo.situ.desc(),Acordo.data_fim,Acordo.nome,Acordo.epe).all() 

        elif lista == 'em execução':

            if type(unid) is str:
                # acordos onde o nome da unidade está no campo unidade_cnpq
                acordos_v = db.session.query(label('id',distinct(Acordo.id)),
                                            Acordo.nome,
                                            Acordo.sei,
                                            Acordo.epe,
                                            Acordo.uf,
                                            Acordo.data_inicio,
                                            Acordo.data_fim,
                                            Acordo.valor_cnpq,
                                            Acordo.valor_epe,
                                            label('unid',Acordo.unidade_cnpq),
                                            Acordo.situ,
                                            Acordo.desc,
                                            cont_prog.c.qtd_prog,
                                            Acordo.siafi,
                                            cont_cham.c.qtd_cha,
                                            Acordo.capital,
                                            Acordo.custeio,
                                            Acordo.bolsas)\
                                    .outerjoin(cont_cham,cont_cham.c.acordo_id == Acordo.id)\
                                    .outerjoin(cont_prog,cont_prog.c.id_acordo == Acordo.id)\
                                    .filter(Acordo.unidade_cnpq.like(unid),
                                            or_(Acordo.situ=='Vigente-Z',Acordo.situ=='Vigente-Esquecido'))\
                                    .order_by(Acordo.data_fim,Acordo.nome,Acordo.epe).all()

            elif type(unid) is list: 
                # acordos onde o nome da unidade está no campo unidade_cnpq
                acordos_v = db.session.query(label('id',distinct(Acordo.id)),
                                             Acordo.nome,
                                             Acordo.sei,
                                             Acordo.epe,
                                             Acordo.uf,
                                             Acordo.data_inicio,
                                             Acordo.data_fim,
                                             Acordo.valor_cnpq,
                                             Acordo.valor_epe,
                                             label('unid',Acordo.unidade_cnpq),
                                             Acordo.situ,
                                             Acordo.desc,
                                             cont_prog.c.qtd_prog,
                                             Acordo.siafi,
                                             cont_cham.c.qtd_cha,
                                            Acordo.capital,
                                            Acordo.custeio,
                                            Acordo.bolsas)\
                                    .outerjoin(cont_cham,cont_cham.c.acordo_id == Acordo.id)\
                                    .outerjoin(cont_prog,cont_prog.c.id_acordo == Acordo.id)\
                                    .filter(Acordo.unidade_cnpq.in_(unid),
                                            or_(Acordo.situ=='Vigente-Z',Acordo.situ=='Vigente-Esquecido'))\
                                    .order_by(Acordo.data_fim,Acordo.nome,Acordo.epe).all() 

        elif lista[:8] == 'programa':

            # pegar coords do programa
            unidades = db.session.query(Programa_CNPq.COORD).filter(Programa_CNPq.COD_PROGRAMA==lista[8:]).all()
            l_unidades = [c.COORD for c in unidades]

            # acordos com nome de unidade no campo unidade_cnpq
            acordos_v = db.session.query(Acordo.id,
                                         Acordo.nome,
                                         Acordo.sei,
                                         Acordo.epe,
                                         Acordo.uf,
                                         Acordo.data_inicio,
                                         Acordo.data_fim,
                                         Acordo.valor_cnpq,
                                         Acordo.valor_epe,
                                         label('unid',Acordo.unidade_cnpq),
                                         Acordo.situ,
                                         Acordo.desc,
                                         cont_prog.c.qtd_prog,
                                         Acordo.siafi,
                                         cont_cham.c.qtd_cha,
                                            Acordo.capital,
                                            Acordo.custeio,
                                            Acordo.bolsas)\
                                  .outerjoin(cont_cham,cont_cham.c.acordo_id == Acordo.id)\
                                  .outerjoin(cont_prog,cont_prog.c.id_acordo == Acordo.id)\
                                  .join(grupo_programa_cnpq, grupo_programa_cnpq.id_acordo == Acordo.id)\
                                  .join(Programa_CNPq, Programa_CNPq.ID_PROGRAMA == grupo_programa_cnpq.id_programa)\
                                  .filter(Programa_CNPq.COD_PROGRAMA == lista[8:],
                                          Acordo.unidade_cnpq.in_(l_unidades))\
                                  .order_by(Acordo.data_fim,Acordo.nome,Acordo.epe).all()

        elif lista[:10] == 'v_programa':

            # acordos com nome de unidade no campo unidade_cnpq
            acordos_v = db.session.query(Acordo.id,
                                         Acordo.nome,
                                         Acordo.sei,
                                         Acordo.epe,
                                         Acordo.uf,
                                         Acordo.data_inicio,
                                         Acordo.data_fim,
                                         Acordo.valor_cnpq,
                                         Acordo.valor_epe,
                                         label('unid',Acordo.unidade_cnpq),
                                         Acordo.situ,
                                         Acordo.desc,
                                         cont_prog.c.qtd_prog,
                                         Acordo.siafi,
                                        cont_cham.c.qtd_cha,
                                            Acordo.capital,
                                            Acordo.custeio,
                                            Acordo.bolsas)\
                                  .outerjoin(cont_cham,cont_cham.c.acordo_id == Acordo.id)\
                                  .outerjoin(cont_prog,cont_prog.c.id_acordo == Acordo.id)\
                                  .join(grupo_programa_cnpq, grupo_programa_cnpq.id_acordo == Acordo.id)\
                                  .join(Programa_CNPq, Programa_CNPq.ID_PROGRAMA == grupo_programa_cnpq.id_programa)\
                                  .filter(Programa_CNPq.COD_PROGRAMA == lista[10:],
                                          or_(Acordo.situ=='Vigente-Z',Acordo.situ=='Esquecido'))\
                                  .order_by(Acordo.data_fim,Acordo.nome,Acordo.epe).all()
   #
        elif lista[:4] == 'edic':
            acordos_v = db.session.query(Acordo.id,
                                       Acordo.nome,
                                       Acordo.sei,
                                       Acordo.epe,
                                       Acordo.uf,
                                       Acordo.data_inicio,
                                       Acordo.data_fim,
                                       Acordo.valor_cnpq,
                                       Acordo.valor_epe,
                                       label('unid',Acordo.unidade_cnpq),
                                       Acordo.situ,
                                       Acordo.desc,
                                       cont_prog.c.qtd_prog,
                                       Acordo.siafi,
                                       cont_cham.c.qtd_cha,
                                            Acordo.capital,
                                            Acordo.custeio,
                                            Acordo.bolsas)\
                                  .outerjoin(cont_cham,cont_cham.c.acordo_id == Acordo.id)\
                                  .outerjoin(cont_prog,cont_prog.c.id_acordo == Acordo.id)\
                                  .filter(Acordo.nome == lista[4:].replace('#$','/'))\
                                  .order_by(Acordo.data_fim,Acordo.nome,Acordo.epe).all()
        
        elif lista[:2] == 'UF':

            if type(unid) is str:
                # acordos onde o nome da unidade está no campo unidade_cnpq
                acordos_v = db.session.query(label('id',distinct(Acordo.id)),
                                            Acordo.nome,
                                            Acordo.sei,
                                            Acordo.epe,
                                            Acordo.uf,
                                            Acordo.data_inicio,
                                            Acordo.data_fim,
                                            Acordo.valor_cnpq,
                                            Acordo.valor_epe,
                                            label('unid',Acordo.unidade_cnpq),
                                            Acordo.situ,
                                            Acordo.desc,
                                            cont_prog.c.qtd_prog,
                                            Acordo.siafi,
                                            cont_cham.c.qtd_cha,
                                            Acordo.capital,
                                            Acordo.custeio,
                                            Acordo.bolsas)\
                                    .outerjoin(cont_cham,cont_cham.c.acordo_id == Acordo.id)\
                                    .outerjoin(cont_prog,cont_prog.c.id_acordo == Acordo.id)\
                                    .filter(Acordo.unidade_cnpq.like(unid),
                                            Acordo.uf == lista[2:4],
                                            or_(Acordo.situ=='Vigente-Z',Acordo.situ=='Vigente-Esquecido'))\
                                    .order_by(Acordo.data_fim,Acordo.nome,Acordo.epe).all()
 
            elif type(unid) is list: 
                # acordos onde o nome da unidade está no campo unidade_cnpq
                acordos_v = db.session.query(label('id',distinct(Acordo.id)),
                                             Acordo.nome,
                                             Acordo.sei,
                                             Acordo.epe,
                                             Acordo.uf,
                                             Acordo.data_inicio,
                                             Acordo.data_fim,
                                             Acordo.valor_cnpq,
                                             Acordo.valor_epe,
                                             label('unid',Acordo.unidade_cnpq),
                                             Acordo.situ,
                                             Acordo.desc,
                                             cont_prog.c.qtd_prog,
                                             Acordo.siafi,
                                             cont_cham.c.qtd_cha,
                                            Acordo.capital,
                                            Acordo.custeio,
                                            Acordo.bolsas)\
                                    .outerjoin(cont_cham,cont_cham.c.acordo_id == Acordo.id)\
                                    .outerjoin(cont_prog,cont_prog.c.id_acordo == Acordo.id)\
                                    .filter(Acordo.unidade_cnpq.in_(unid),
                                            Acordo.uf == lista[2:4],
                                            or_(Acordo.situ=='Vigente-Z',Acordo.situ=='Vigente-Esquecido'))\
                                    .order_by(Acordo.data_fim,Acordo.nome,Acordo.epe).all() 

        elif lista[:7] == 'PROG_UF':

            if type(unid) is str:
                # acordos onde o nome da unidade está no campo unidade_cnpq
                acordos_v = db.session.query(label('id',distinct(Acordo.id)),
                                            Acordo.nome,
                                            Acordo.sei,
                                            Acordo.epe,
                                            Acordo.uf,
                                            Acordo.data_inicio,
                                            Acordo.data_fim,
                                            Acordo.valor_cnpq,
                                            Acordo.valor_epe,
                                            label('unid',Acordo.unidade_cnpq),
                                            Acordo.situ,
                                            Acordo.desc,
                                            cont_prog.c.qtd_prog,
                                            Acordo.siafi,
                                            cont_cham.c.qtd_cha,
                                            Acordo.capital,
                                            Acordo.custeio,
                                            Acordo.bolsas)\
                                    .outerjoin(cont_cham,cont_cham.c.acordo_id == Acordo.id)\
                                    .outerjoin(cont_prog,cont_prog.c.id_acordo == Acordo.id)\
                                    .join(grupo_programa_cnpq, grupo_programa_cnpq.id_acordo == Acordo.id)\
                                    .join(Programa_CNPq, Programa_CNPq.ID_PROGRAMA == grupo_programa_cnpq.id_programa)\
                                    .filter(Acordo.unidade_cnpq.like(unid),
                                            Acordo.uf == lista[7:9],
                                            Programa_CNPq.COD_PROGRAMA == lista[9:22],
                                            or_(Acordo.situ=='Vigente-Z',Acordo.situ=='Vigente-Esquecido'))\
                                    .order_by(Acordo.data_fim,Acordo.nome,Acordo.epe).all()
 
            elif type(unid) is list: 
                # acordos onde o nome da unidade está no campo unidade_cnpq
                acordos_v = db.session.query(label('id',distinct(Acordo.id)),
                                             Acordo.nome,
                                             Acordo.sei,
                                             Acordo.epe,
                                             Acordo.uf,
                                             Acordo.data_inicio,
                                             Acordo.data_fim,
                                             Acordo.valor_cnpq,
                                             Acordo.valor_epe,
                                             label('unid',Acordo.unidade_cnpq),
                                             Acordo.situ,
                                             Acordo.desc,
                                             cont_prog.c.qtd_prog,
                                             Acordo.siafi,
                                            cont_cham.c.qtd_cha,
                                            Acordo.capital,
                                            Acordo.custeio,
                                            Acordo.bolsas)\
                                    .outerjoin(cont_cham,cont_cham.c.acordo_id == Acordo.id)\
                                    .outerjoin(cont_prog,cont_prog.c.id_acordo == Acordo.id)\
                                    .join(grupo_programa_cnpq, grupo_programa_cnpq.id_acordo == Acordo.id)\
                                    .join(Programa_CNPq, Programa_CNPq.ID_PROGRAMA == grupo_programa_cnpq.id_programa)\
                                    .filter(Acordo.unidade_cnpq.in_(unid),
                                            Acordo.uf == lista[7:9],
                                            Programa_CNPq.COD_PROGRAMA == lista[9:22],
                                            or_(Acordo.situ=='Vigente-Z',Acordo.situ=='Vigente-Esquecido'))\
                                    .order_by(Acordo.data_fim,Acordo.nome,Acordo.epe).all() 

        quantidade = len(acordos_v)

        acordos = []

        for acordo in acordos_v:

            if acordo.data_fim:
                dias = (acordo.data_fim - hoje).days
            else:
                dias = 999

            valor_global = acordo.valor_epe + acordo.valor_cnpq
            valor_cnpq   = acordo.valor_cnpq
            valor_bolsas = acordo.bolsas
            # valor_cnpq   = locale.currency(acordo.valor_cnpq, symbol=False, grouping = True)
            valor_epe    = locale.currency(acordo.valor_epe, symbol=False, grouping = True)

            # pega quantidade de mães, filhos do acordo e totaliza o que foi pago
            procs_mae = db.session.query(Acordo_ProcMae.proc_mae_id,
                                         Processo_Mae.proc_mae,
                                         Processo_Mae.situ_mae,
                                         Processo_Mae.pago_capital,
                                         Processo_Mae.pago_custeio)\
                                  .join(Processo_Mae, Processo_Mae.id == Acordo_ProcMae.proc_mae_id)\
                                  .filter(Acordo_ProcMae.acordo_id == acordo.id)\
                                  .all()
            qtd_proc_mae = len(procs_mae)

            qtd_filhos_acordo = 0
            pago_acordo = 0
            a_pagar = 0

            for proc in procs_mae:

                if proc.pago_capital:
                    pago_acordo += proc.pago_capital
                if proc.pago_custeio:
                    pago_acordo += proc.pago_custeio
 
                filhos = db.session.query(Processo_Filho.proc_mae,
                                          label('qtd_filhos',func.count(Processo_Filho.processo)),
                                          label('pago_filhos',func.sum(Processo_Filho.pago_total)),
                                          label('a_pagar_filhos',func.sum(Processo_Filho.valor_apagar)))\
                                   .filter(Processo_Filho.proc_mae == proc.proc_mae)\
                                   .group_by(Processo_Filho.proc_mae)\
                                   .first()
                if filhos:                   
                    qtd_filhos_acordo += int(filhos.qtd_filhos)
                    pago_acordo += filhos.pago_filhos
                    if filhos.a_pagar_filhos != None:
                        a_pagar += filhos.a_pagar_filhos

            # ver como receber valores pagos em capital e custeio para abater no calculo do saldo
            # pago_capital = ....
            # pago_custeio = ....

            #
            # verifica ser o acordo tem demandas de indicação de bolsisstas
            indic = 0
            indic = db.session.query(Demanda.tipo).filter(Demanda.sei == acordo.sei, Demanda.tipo == 'Bolsistas - Indicação na PICC').count()

            #verificar se situação do acordo pode ser modificada

            situ = acordo.situ
            alterar_sit = False

            if acordo.data_fim != None and acordo.data_fim != '' and acordo.data_inicio != None and acordo.data_inicio != '':
                if situ != 'Vigente-Z' and qtd_proc_mae > 0 and acordo.data_fim >= hoje:
                    situ = 'Vigente-Z'
                    alterar_sit = True
                if situ[0:8] != 'Expirado' and situ != 'Não executado' and acordo.data_fim < hoje:
                    situ = 'Expirado (sem RTF)'
                    alterar_sit = True
                if situ == 'Expirado' and acordo.data_fim < hoje:
                    situ = 'Expirado (sem RTF)'
                    alterar_sit = True
                if (situ[0:8] == 'Expirado' or situ == 'Vigente-Z' or situ == 'Preparação' or situ == 'Vigente-Esquecido' or situ == '' or situ == 'Consta indicação de bolsista') and\
                            qtd_proc_mae == 0 and acordo.data_fim > hoje and indic == 0 and (acordo.data_inicio+datetime.timedelta(days=90)) >= hoje:
                    situ = 'Assinado'
                    alterar_sit = True
                if (situ[0:8] == 'Expirado' or situ == 'Vigente-Z' or situ == 'Assinado' or situ == 'Preparação' or situ == '' or situ == 'Consta indicação de bolsista') and\
                            qtd_proc_mae == 0 and acordo.data_fim > hoje and indic == 0 and (acordo.data_inicio+datetime.timedelta(days=90)) < hoje:
                    situ = 'Vigente-Esquecido'
                    alterar_sit = True
                if (situ[0:8] == 'Assinado' or situ == 'Vigente-Esquecido' or situ == 'Aguarda Folha' or situ == 'Bolsas foram indicadas') and indic > 0:
                    situ = 'Consta indicação de bolsista'
                    alterar_sit = True
            else:
                if situ != 'Preparação' and situ != "Não executado":
                    situ = 'Preparação'
                    alterar_sit = True

            if alterar_sit:
                acordo_alterar_sit = Acordo.query.get_or_404(acordo.id)
                acordo_alterar_sit.situ = situ
                db.session.commit()
            #
            acordos.append([acordo.id,
                            '',
                            acordo.nome, 
                            acordo.sei, 
                            acordo.epe, 
                            acordo.uf,
                            acordo.data_inicio,
                            acordo.data_fim,
                            valor_epe,
                            valor_cnpq,
                            qtd_proc_mae,
                            qtd_filhos_acordo,
                            pago_acordo,
                            0,
                            0,
                            acordo.unid,
                            dias,
                            '',
                            0,
                            situ,
                            acordo.desc,
                            acordo.qtd_prog,
                            acordo.siafi,
                            acordo.qtd_cha,
                            valor_global,
                            a_pagar,
                            valor_bolsas])

        try:
            cria_csv('/app/project/static/acordos.csv',
                    ['id','***','nome','sei','epe','uf','ini','fim','valor_epe','valor_cnpq','qtd_proc_mae','qtd_filhos','pago','a_pagar','saldo',\
                    'coord','dias','***','qtd_cpfs','situ','desc','qtd_prog','siafi','valor_global','a_pagar','valor_bolsas'],
                    acordos)
            tem_csv = True          
        except:
            flash('Arquivo csv não foi criado!','erro') 
            tem_csv = False             

        # o comandinho mágico que permite fazer o download de um arquivo
        if tem_csv == True:
            send_from_directory('/app/project/static', 'acordos.csv')                     

        return render_template('lista_acordos.html', 
                               acordos=acordos,
                               quantidade=quantidade,
                               lista=lista,
                               form=form,
                               data_cha = data_cha,
                               tem_csv = tem_csv)


### VISUALIZAR E ATUALIZAR detalhes de Acordo

@acordos.route("/<int:acordo_id>/<lista>/update", methods=['GET', 'POST'])
@login_required
def update(acordo_id,lista):
    """
    +---------------------------------------------------------------------------------------+
    |Permite ver e atualizar os dados de um acordo selecionado na tela de consulta.         |
    |                                                                                       |
    |Recebe o ID do acordo como parâmetro.                                                  |
    +---------------------------------------------------------------------------------------+
    """
    unidade = current_user.coord

    # se unidade for pai, junta ela com seus filhos
    hierarquia = db.session.query(Coords.sigla).filter(Coords.pai == unidade).all()

    if hierarquia:
        l_unid = [f.sigla for f in hierarquia]
        l_unid.append(unidade)
    else:
        l_unid = [unidade]

    lista_coords = [(c,c) for c in l_unid]
    lista_coords.insert(0,('',''))

    chamadas_s = []

    acordo = Acordo.query.get_or_404(acordo_id)

    # contabiliza quantidade de programas do acordo
    cont_prog = db.session.query(grupo_programa_cnpq.id_acordo,
                                label('qtd_prog',func.count(grupo_programa_cnpq.id_programa)))\
                            .filter(grupo_programa_cnpq.id_acordo == acordo_id)\
                            .group_by(grupo_programa_cnpq.id_acordo)\
                            .first()
    if cont_prog:
        qtd_prog = cont_prog.qtd_prog
    else:
        qtd_prog = 0                        

    # contabiliza quantidade de chamadas PICC do acordo
    cont_cham = db.session.query(chamadas_cnpq_acordos.acordo_id,
                                    label('qtd_cha',func.count(chamadas_cnpq_acordos.id)))\
                            .filter(chamadas_cnpq_acordos.acordo_id == acordo_id)\
                            .group_by(chamadas_cnpq_acordos.acordo_id)\
                            .first()
    if cont_cham:
        qtd_cha = cont_cham.qtd_cha
    else:
        qtd_cha = 0  

    # pega quantidade de mães, filhos do acordo e totaliza o que foi pago
    procs_mae = db.session.query(Acordo_ProcMae.proc_mae_id,
                                 Processo_Mae.proc_mae,
                                 Processo_Mae.situ_mae,
                                 Processo_Mae.pago_capital,
                                 Processo_Mae.pago_custeio)\
                            .join(Processo_Mae, Processo_Mae.id == Acordo_ProcMae.proc_mae_id)\
                            .filter(Acordo_ProcMae.acordo_id == acordo.id)\
                            .all()
    qtd_proc_mae = len(procs_mae)

    qtd_filhos_acordo = 0
    pago_capital = 0
    pago_custeio = 0
    pago_bolsas  = 0

    for proc in procs_mae:

        if proc.pago_capital:
            pago_capital += proc.pago_capital
        if proc.pago_custeio:
            pago_custeio += proc.pago_custeio

        filhos = db.session.query(Processo_Filho.proc_mae,
                                    label('qtd_filhos',func.count(Processo_Filho.processo)),
                                    label('pago_filhos',func.sum(Processo_Filho.pago_total)))\
                            .filter(Processo_Filho.proc_mae == proc.proc_mae)\
                            .group_by(Processo_Filho.proc_mae)\
                            .first()
        if filhos:                   
            qtd_filhos_acordo += int(filhos.qtd_filhos)
            pago_bolsas += filhos.pago_filhos

    chamadas = db.session.query(Chamadas.id,
                                Chamadas.chamada,
                                Chamadas.qtd_projetos,
                                Chamadas.vl_total_chamada,
                                Chamadas.doc_sei,
                                Chamadas.obs,
                                Chamadas.id_relaciona)\
                         .filter(Chamadas.id_relaciona == str(acordo.id)).all()
    qtd_chamadas = len(chamadas)                            

    chamadas_s = []
    chamadas_tot = 0
    total_projetos = 0

    for chamada in chamadas:
        chamadas_s.append([chamada.id, 
                           chamada.chamada,
                           chamada.qtd_projetos,
                           locale.currency(chamada.vl_total_chamada, symbol=False, grouping = True),
                           chamada.doc_sei, 
                           chamada.obs, 
                           chamada.id_relaciona])
        chamadas_tot += chamada.vl_total_chamada
        total_projetos += chamada.qtd_projetos
    
    try:
        sei = str(acordo.sei).split('/')[0]+'_'+str(acordo.sei).split('/')[1]
    except:
        sei = acordo.sei

    form = AcordoForm()    

    form.unid.choices = lista_coords

    form.situacao.choices=[('',''),
                       ('Preparação','Preparação'),
                       ('Assinado','Assinado'),
                       ('Vigente-Esquecido','Vigente-Esquecido'),
                       ('Aguarda Folha','Aguarda Folha'),
                       ('Vigente-Z','Vigente'),
                       ('Expirado (sem RTF)','Expirado (sem RTF)'),
                       ('Expirado (RTF aguarda análise)','Expirado (RTF aguarda análise)'),
                       ('Expirado (RTF aprovado, mas há processo(s) à finalizar)','Expirado (RTF aprovado, mas há processo(s) à finalizar)'),
                       ('Expirado (RTF APROVADO!)','Expirado (RTF APROVADO!)'),
                       ('Expirado (sit. 71 mãe(s), mas há filho(s) pendente(s))','Expirado (sit. 71 mãe(s), mas há filho(s) pendente(s))'),
                       ('Expirado (sit. 71 mãe(s) e filho(s))','Expirado (sit. 71 mãe(s) e filho(s))'),
                       ('Não executado','Não executado')]

    if form.validate_on_submit():

        valor_cnpq = float(form.valor_cnpq.data.replace('.','').replace(',','.'))
        valor_epe  = float(form.valor_epe.data.replace('.','').replace(',','.'))
        capital    = float(form.capital.data.replace('.','').replace(',','.'))
        custeio    = float(form.custeio.data.replace('.','').replace(',','.'))
        bolsas     = float(form.bolsas.data.replace('.','').replace(',','.'))

        valor = valor_cnpq + valor_epe
        nds = capital + custeio + bolsas

        if round(nds,2) != round(valor,2) and (capital > 0 or custeio > 0 or bolsas > 0):
            flash('Atenção: Soma Capital, Custeio e Bolsas não corresponde à soma dos valores do acordo/TED \
                  (Aporte: '+str(locale.currency(round(valor,2), symbol=False, grouping = True ))+', \
                   Soma NDs: '+str(locale.currency(round(nds,2), symbol=False, grouping = True ))+')!','perigo')
            # return redirect(url_for('acordos.update', acordo_id=acordo_id, lista=lista))

        acordo.nome         = form.nome.data
        acordo.sei          = form.sei.data
        acordo.epe          = form.epe.data
        acordo.uf           = form.uf.data
        acordo.data_inicio  = form.data_inicio.data
        acordo.data_fim     = form.data_fim.data
        acordo.valor_cnpq   = valor_cnpq
        acordo.valor_epe    = valor_epe
        acordo.unidade_cnpq = form.unid.data
        acordo.situ         = form.situacao.data
        acordo.desc         = form.desc.data
        acordo.capital      = capital
        acordo.custeio      = custeio
        acordo.bolsas       = bolsas
        acordo.siafi        = form.siafi.data

        db.session.commit()

        # atualiza demandas associadas em caso de alteração do SEI do acordo.
        if form.sei.data != acordo.sei:
            demandas = db.session.query(Demanda).filter(Demanda.sei == acordo.sei).all()
            for demanda in demandas:
                demanda.sei = form.sei.data
            db.session.commit()

        registra_log_auto(current_user.id,None,'aco')

        flash('Acordo atualizado!','sucesso')
        return redirect(url_for('acordos.update', acordo_id=acordo_id, lista=lista))

    # traz a informação atual do acordo
    form.nome.data        = acordo.nome
    form.desc.data        = acordo.desc
    form.sei.data         = acordo.sei
    form.epe.data         = acordo.epe
    form.uf.data          = acordo.uf
    form.data_inicio.data = acordo.data_inicio
    form.data_fim.data    = acordo.data_fim
    form.valor_cnpq.data  = locale.currency( acordo.valor_cnpq, symbol=False, grouping = True )
    form.valor_epe.data   = locale.currency( acordo.valor_epe, symbol=False, grouping = True )
    if acordo.unidade_cnpq.isdigit():
        form.unid.data = None 
    else:
        form.unid.data = acordo.unidade_cnpq  
    form.situacao.data     = acordo.situ
    print ('** acordo.situ: ',acordo.situ, type(acordo.situ))
    print ('** form.situacao.data: ', form.situacao.data, type(form.situacao.data))
    form.capital.data  = locale.currency( acordo.capital, symbol=False, grouping = True )
    form.custeio.data  = locale.currency( acordo.custeio, symbol=False, grouping = True )
    form.bolsas.data   = locale.currency( acordo.bolsas, symbol=False, grouping = True )
    form.siafi.data    = acordo.siafi


    return render_template('add_acordo.html', title='Update',
                            chamadas=chamadas_s,
                            qtd_chamadas=qtd_chamadas,
                            qtd_proj = total_projetos,
                            chamadas_tot=locale.currency(chamadas_tot, symbol=False, grouping = True),
                            acordo_id=acordo_id,
                            sei=sei,
                            prog="",
                            edic=acordo.nome,
                            epe=acordo.epe,
                            uf=acordo.uf,
                            procs_mae=procs_mae,
                            qtd_procs_mae=len(procs_mae),
                            form=form,
                            cont_prog=qtd_prog,
                            cont_cham=qtd_cha,
                            pago_capital=pago_capital,
                            pago_custeio=pago_custeio,
                            pago_bolsas=pago_bolsas)

# lista acordo de associado a um processo-mãe
@acordos.route("/<int:proc_mae_id>/consulta_acordo_proc_mae")
@login_required
def consulta_acordo_proc_mae(proc_mae_id):
    """
    +---------------------------------------------------------------------------------------+
    |Mostra o acordo associado a u processo-mãe informado.                                  |
    |                                                                                       |
    |Recebe o id do processo mãe como parâmetro.                                            |
    +---------------------------------------------------------------------------------------+
    """

    acordo = db.session.query(Acordo_ProcMae)\
                       .filter(Acordo_ProcMae.proc_mae_id == proc_mae_id)\
                       .first()


    return redirect(url_for('acordos.update', acordo_id=acordo.acordo_id, lista='todos'))


@acordos.route("/<int:acordo_id>/chamadas_acordo", methods=['GET', 'POST'])
@login_required
def chamadas_acordo(acordo_id):
    """
    +---------------------------------------------------------------------------------------+
    |Lista das chamadas do CNPq de um acordo.                                               |
    |                                                                                       |
    |Recebe o ID do acordo como parâmetro.                                                  |
    +---------------------------------------------------------------------------------------+
    """

    acordo = Acordo.query.get_or_404(acordo_id)

    processos = db.session.query(Processo_Mae.id_chamada,
                                 label('qtd_proc',func.count(Processo_Mae.id)))\
                          .group_by(Processo_Mae.id_chamada)\
                          .subquery()

    chamadas = db.session.query(chamadas_cnpq.id,
                                chamadas_cnpq.tipo,
                                chamadas_cnpq.nome,
                                chamadas_cnpq.sigla,
                                chamadas_cnpq.valor,
                                chamadas_cnpq.cod_programa,
                                chamadas_cnpq.id_dw,
                                chamadas_cnpq.qtd_processos,
                                processos.c.qtd_proc)\
                         .join(chamadas_cnpq_acordos, chamadas_cnpq_acordos.chamada_cnpq_id == chamadas_cnpq.id)\
                         .outerjoin(processos, processos.c.id_chamada == chamadas_cnpq.id)\
                         .filter(chamadas_cnpq_acordos.acordo_id == acordo.id)\
                         .all()

    qtd_chamadas = len(chamadas)                            

    chamadas_tot = 0
    total_projetos = 0
    total_processos = 0

    for chamada in chamadas:
        chamadas_tot += chamada.valor
        if chamada.qtd_processos:
            total_processos += chamada.qtd_processos
    
    return render_template('lista_chamadas_acordo.html',
                            acordo = acordo,
                            chamadas = chamadas,
                            qtd_chamadas = qtd_chamadas,
                            chamadas_tot = chamadas_tot,
                            total_processos = total_processos)

### LISTAR processos de uma chamada do CNPq

@acordos.route("/<int:chamada_id_dw>/processos_chamada")
def processos_chamada(chamada_id_dw):
    """
    +---------------------------------------------------------------------------------------+
    |Lista processos de uma chamda do CNPq obtidos via carga de dados proveniente do DW.    |
    +---------------------------------------------------------------------------------------+
    """

    chamada = db.session.query(chamadas_cnpq)\
                        .filter(chamadas_cnpq.id_dw == chamada_id_dw).first()

    processos = db.session.query(Processo_Mae.id,
                                 Acordo_ProcMae.proc_mae_id,
                                 Processo_Mae.proc_mae,
                                 Processo_Mae.coordenador,
                                 Processo_Mae.inic_mae,
                                 Processo_Mae.term_mae,
                                 Processo_Mae.situ_mae,
                                 label('qtd_filhos',func.count(distinct(Processo_Filho.processo))),
                                 label('pago',func.sum(Processo_Filho.pago_total)),
                                 label('max_ult_pag',func.max(Processo_Filho.dt_ult_pag)),
                                 Processo_Mae.pago_custeio,
                                 Processo_Mae.pago_capital,
                                 label('pago_cap_cus',Processo_Mae.pago_custeio + Processo_Mae.pago_capital))\
                          .outerjoin(Processo_Filho, Processo_Filho.proc_mae == Processo_Mae.proc_mae)\
                          .outerjoin(Acordo_ProcMae, Acordo_ProcMae.proc_mae_id == Processo_Mae.id)\
                          .filter(Processo_Mae.id_chamada == chamada.id)\
                          .group_by(Processo_Mae.id,
                                    Acordo_ProcMae.proc_mae_id,
                                    Processo_Mae.proc_mae,
                                    Processo_Mae.coordenador,
                                    Processo_Mae.inic_mae,
                                    Processo_Mae.term_mae,
                                    Processo_Mae.situ_mae,
                                    Processo_Mae.pago_custeio,
                                    Processo_Mae.pago_capital)\
                          .all()                                           

    qtd_processos = len(processos)

    if qtd_processos == 0:
        flash('Para obter lista dos processos desta chamada, será necessário efetuar nova carga de chamadas!','perigo')


    return render_template('lista_processos_mae.html',procs_mae=processos,
                                                      qtd_processos=qtd_processos,
                                                      chamada = chamada.nome,
                                                      acordo_id=None,
                                                      acordo_tit=None)

### associar programa a Acordo

@acordos.route("/programa_acordo/<int:id_acordo>", methods=['GET', 'POST'])
@login_required
def programa_acordo(id_acordo):
    """
    +---------------------------------------------------------------------------------------+
    |Permite associar programas a um acordo.                                                |
    +---------------------------------------------------------------------------------------+
    """
    unidade = current_user.coord

    # se unidade for pai, junta ela com seus filhos
    hierarquia = db.session.query(Coords.sigla).filter(Coords.pai == unidade).all()

    if hierarquia:
        l_unid = [f.sigla for f in hierarquia]
        l_unid.append(unidade)
    else:
        l_unid = [unidade]


    acordo = db.session.query(Acordo).filter(Acordo.id==id_acordo).first()

    programas_cnpq = db.session.query(Programa_CNPq.ID_PROGRAMA,
                                      Programa_CNPq.COD_PROGRAMA,
                                      Programa_CNPq.SIGLA_PROGRAMA,
                                      Programa_CNPq.COORD)\
                               .filter(Programa_CNPq.COORD.in_(l_unid))\
                               .order_by(Programa_CNPq.NOME_PROGRAMA)\
                               .all()
    lista_progs = [(str(prog.ID_PROGRAMA), prog.COD_PROGRAMA +' - '+ prog.SIGLA_PROGRAMA +' - '+prog.COORD) for prog in programas_cnpq]
    lista_progs.insert(0,('',''))

    form = ProgAcordoForm()

    form.programa_cnpq.choices = lista_progs

    if form.validate_on_submit():

        for p in form.programa_cnpq.data:
            novo_grupo = grupo_programa_cnpq(id_acordo   = id_acordo,
                                             id_programa = int(p),
                                             cod_programa = None)
            db.session.add(novo_grupo)                                 
        db.session.commit()

        flash('Programa(s) associados ao Acordo!','sucesso')
        return redirect(url_for('acordos.lista_acordos',lista='todos',coord='usu'))

    return render_template('add_programa_acordo.html',
                            acordo=acordo,
                            form=form)        


### associar chamada a Acordo

@acordos.route("/associa_chamada/<int:id_acordo>", methods=['GET', 'POST'])
@login_required
def associa_chamada(id_acordo):
    """
    +---------------------------------------------------------------------------------------+
    |Permite associar uma chamada presente em uma lista a um acordo.                        |
    +---------------------------------------------------------------------------------------+
    """

    acordo = Acordo.query.get_or_404(id_acordo)

    programas = db.session.query(Programa_CNPq.COD_PROGRAMA)\
                          .join(grupo_programa_cnpq, grupo_programa_cnpq.id_programa == Programa_CNPq.ID_PROGRAMA)\
                          .filter(grupo_programa_cnpq.id_acordo == acordo.id)\
                          .all()

    l_programas = set([p.COD_PROGRAMA for p in programas])

    chamadas = db.session.query(chamadas_cnpq.id,
                                chamadas_cnpq.tipo,
                                chamadas_cnpq.nome,
                                chamadas_cnpq.sigla)\
                         .filter(chamadas_cnpq.cod_programa.in_(l_programas))\
                         .order_by(chamadas_cnpq.nome)\
                         .all()

    lista_chamadas = [(str(c.id),(c.tipo +' - '+ c.sigla +' - '+ c.nome)[:125]+'...') if len((c.tipo +' - '+ c.sigla +' - '+ c.nome)) >125 else (str(c.id),(c.tipo +' - '+ c.sigla +' - '+ c.nome)) for c in chamadas]
    lista_chamadas.insert(0,('',''))

    form = ChamadaAcordoForm()

    form.chamada.choices= lista_chamadas

    if form.validate_on_submit():

        for c in form.chamada.data:

            chamada_cnpq_acordo = chamadas_cnpq_acordos (acordo_id       = id_acordo,
                                                         chamada_cnpq_id = c)
            db.session.add(chamada_cnpq_acordo)

        db.session.commit()

        flash('Chamada(s) associada(s) ao Acordo!','sucesso')

        # return redirect(url_for('acordos.chamadas_acordo', acordo_id=id_acordo))
        return redirect(url_for('acordos.update', acordo_id=id_acordo, lista='todos'))


    return render_template('add_chamada_acordo.html',
                            acordo=acordo,
                            form=form)   


### desassociar chamada de Acordo

@acordos.route("<int:id>/<int:id_acordo>/desassocia_chamada", methods=['GET', 'POST'])
@login_required
def desassocia_chamada(id,id_acordo):
    """
    +---------------------------------------------------------------------------------------+
    |Permite desassociar uma chamada de um acordo, bem como processos_mae envolvidos.       |
    +---------------------------------------------------------------------------------------+
    """

    # remover associações dos processos mãe

    procs_mae = db.session.query(Processo_Mae).filter(Processo_Mae.id_chamada == id).all()

    for proc in procs_mae:

        proc_acordo = db.session.query(Acordo_ProcMae)\
                                .filter(Acordo_ProcMae.proc_mae_id == proc.id,
                                        Acordo_ProcMae.acordo_id == id_acordo)\
                                .delete()

    db.session.commit()    

    # remover associação da chamada com o acordo

    chamadas_desassocia = db.session.query(chamadas_cnpq_acordos)\
                                    .filter(chamadas_cnpq_acordos.acordo_id == id_acordo,
                                           chamadas_cnpq_acordos.chamada_cnpq_id == id)\
                                    .delete()
   
    db.session.commit()

    flash('Chamada desassociada do Acordo, bem como processos-mãe relacionados!','sucesso')

    return redirect(url_for('acordos.lista_acordos',lista='em execução',coord='usu'))
 

### CRIAR Acordo

@acordos.route("/criar", methods=['GET', 'POST'])
@login_required
def cria_acordo():
    """
    +---------------------------------------------------------------------------------------+
    |Permite registrar os dados de um acordo.                                               |
    +---------------------------------------------------------------------------------------+
    """
    unidade = current_user.coord

    # se unidade for pai, junta ela com seus filhos
    hierarquia = db.session.query(Coords.sigla).filter(Coords.pai == unidade).all()

    if hierarquia:
        l_unid = [f.sigla for f in hierarquia]
        l_unid.append(unidade)
    else:
        l_unid = [unidade]

    lista_coords = [(c,c) for c in l_unid]
    lista_coords.insert(0,('',''))

    form = AcordoForm()

    form.unid.choices = lista_coords

    if form.validate_on_submit():

        valor_cnpq    = float(form.valor_cnpq.data.replace('.','').replace(',','.'))
        valor_epe     = float(form.valor_epe.data.replace('.','').replace(',','.'))
        capital       = float(form.capital.data.replace('.','').replace(',','.'))
        custeio       = float(form.custeio.data.replace('.','').replace(',','.'))
        bolsas        = float(form.bolsas.data.replace('.','').replace(',','.'))

        valor = valor_cnpq + valor_epe
        nds = capital + custeio + bolsas


        if round(nds,2) != round(valor,2) and (capital != 0 or custeio != 0 or bolsas != 0):
            flash('Atenção: Soma Capital, Custeio e Bolsas não corresponde à soma dos valores do acordo/TED \
                  (Aporte: '+str(locale.currency(round(valor,2), symbol=False, grouping = True ))+', \
                   Soma NDs: '+str(locale.currency(round(nds,2), symbol=False, grouping = True ))+')!','perigo')
            # return redirect(url_for('acordos.cria_acordo'))

        acordo = Acordo(nome          = form.nome.data,
                        desc          = form.desc.data,
                        sei           = form.sei.data,
                        epe           = form.epe.data,
                        uf            = form.uf.data,
                        data_inicio   = form.data_inicio.data,
                        data_fim      = form.data_fim.data,
                        valor_cnpq    = valor_cnpq,
                        valor_epe     = valor_cnpq,
                        unidade_cnpq = form.unid.data,
                        situ          = form.situacao.data,
                        capital       = capital,
                        custeio       = custeio,
                        bolsas        = bolsas,
                        siafi         = form.siafi.data)

        db.session.add(acordo)
        db.session.commit()

        registra_log_auto(current_user.id,None,'aco')

        flash('Acordo criado!','sucesso')
        return redirect(url_for('acordos.lista_acordos',lista='todos',coord='usu'))


    return render_template('add_acordo.html',
                            acordo_id=0,
                            form=form)

## Deletar um acordo

@acordos.route("/<int:acordo_id>/deleta_acordo", methods=['GET', 'POST'])
@login_required
def deleta_acordo(acordo_id):
    """
    +---------------------------------------------------------------------------------------+
    |Permite deletar o registro de um acordo.                                               |
    +---------------------------------------------------------------------------------------+
    """

    acordo = Acordo.query.get_or_404(acordo_id)

    db.session.delete(acordo)
    db.session.commit()

    ## deletar associações do acordo com programas
    programas = db.session.query(grupo_programa_cnpq).filter(grupo_programa_cnpq.id_acordo==acordo_id).all() 
    db.session.delete(programas)
    db.session.commit()

    ## deletar associações do acordo com capital_custeio
    cap_cus = db.session.query(capital_custeio).filter(capital_custeio.id_acordo==acordo_id).all() 
    db.session.delete(cap_cus)
    db.session.commit()

    registra_log_auto(current_user.id,None,'ade')

    flash ('Acordo deletado!','sucesso')

    return redirect(url_for('acordos.lista_acordos',lista='todos',coord='usu'))


# lista das demandas relacionadas a um acordo

@acordos.route('/<acordo_id>/acordo_demandas')
def acordo_demandas (acordo_id):
    """+--------------------------------------------------------------------------------------+
       |Mostra as demandas relacionadas a um acordo quando seu NUP é selecionado em uma       |
       |lista de acordos.                                                                     |
       |Recebe o id do acordo como parâmetro.                                                 |
       +--------------------------------------------------------------------------------------+
    """

    acordo_SEI = db.session.query(Acordo.sei,Acordo.nome).filter_by(id=acordo_id).first()

    SEI = acordo_SEI.sei
    SEI_s = str(SEI).split('/')[0]+'_'+str(SEI).split('/')[1]

    demandas_count = Demanda.query.filter(Demanda.sei.like('%'+SEI+'%')).count()

    demandas = Demanda.query.filter(Demanda.sei.like('%'+SEI+'%'))\
                            .order_by(Demanda.data.desc()).all()

    autores=[]
    for demanda in demandas:
        autores.append(str(User.query.filter_by(id=demanda.user_id).first()).split(';')[0])

    dados = [acordo_SEI.nome,SEI_s,'0','0']

    return render_template('SEI_demandas.html',demandas_count=demandas_count,demandas=demandas,sei=SEI, autores=autores,dados=dados)

#
### CRIAR programa do CNPq

@acordos.route("/cria_programa_cnpq", methods=['GET', 'POST'])
@login_required
def cria_programa_cnpq():
    """
    +---------------------------------------------------------------------------------------+
    |Permite registrar os dados de um programa do CNPq.                                     |
    +---------------------------------------------------------------------------------------+
    """

    unidade = current_user.coord

    # se unidade for pai, junta ela com seus filhos
    hierarquia = db.session.query(Coords.sigla).filter(Coords.pai == unidade).all()

    if hierarquia:
        l_unid = [f.sigla for f in hierarquia]
        l_unid.append(unidade)
    else:
        l_unid = [unidade]

    lista_coords = [(c,c) for c in l_unid]
    lista_coords.insert(0,('',''))

    form = Programa_CNPqForm()
    form.coord.choices = lista_coords

    if form.validate_on_submit():

        # last_programa_cnpq = db.session.query(Programa_CNPq).order_by(Programa_CNPq.ID_PROGRAMA.desc()).first()

        programa_cnpq = Programa_CNPq(COD_PROGRAMA   = form.cod_programa.data,
                                      NOME_PROGRAMA  = form.nome_programa.data,
                                      SIGLA_PROGRAMA = form.sigla_programa.data,
                                      COORD          = form.coord.data)

        db.session.add(programa_cnpq)
        db.session.commit()

        registra_log_auto(current_user.id,None,'pac')

        flash('Programa do CNPq registrado!','sucesso')
        return redirect(url_for('acordos.lista_programa_cnpq'))

    return render_template('cria_programa_cnpq.html', form=form)

#
#
### LISTAR programas do CNPq

@acordos.route("/lista_programa_cnpq")
def lista_programa_cnpq():
    """
    +---------------------------------------------------------------------------------------+
    |Permite listar os programas do CNPq vinculados à unidade do usuário logado.            |
    +---------------------------------------------------------------------------------------+
    """

    unidade = current_user.coord

    # se unidade for pai, junta ela com seus filhos
    hierarquia = db.session.query(Coords.sigla).filter(Coords.pai == unidade).all()

    if hierarquia:
        l_unid = [f.sigla for f in hierarquia]
        l_unid.append(unidade)
    else:
        l_unid = [unidade]

    programas = db.session.query(Programa_CNPq)\
                          .filter(Programa_CNPq.COORD.in_(l_unid))\
                          .all()
    
    quantidade = len(programas)

    return render_template('lista_programa_cnpq.html',programas=programas, quantidade=quantidade)


# lista programas de um acordo
@acordos.route("/<int:id_acordo>/lista_programas_acordo")
@login_required
def lista_programas_acordo(id_acordo):
    """
    +---------------------------------------------------------------------------------------+
    |Lista programas que o acordo usa para efeturar pagamentos.                             |
    |                                                                                       |
    |Recebe o ID do acordo como parâmetro.                                                  |
    +---------------------------------------------------------------------------------------+
    """

    lista_programas = db.session.query(grupo_programa_cnpq.id_acordo,
                                       Programa_CNPq.COD_PROGRAMA,
                                       Programa_CNPq.SIGLA_PROGRAMA,
                                       Programa_CNPq.NOME_PROGRAMA)\
                                .join(Programa_CNPq, Programa_CNPq.ID_PROGRAMA == grupo_programa_cnpq.id_programa)\
                                .filter(grupo_programa_cnpq.id_acordo == id_acordo)\
                                .order_by(Programa_CNPq.SIGLA_PROGRAMA)\
                                .all()
    qtd_progs = len(lista_programas)

    # pega nome do acordo
    acordo = db.session.query(Acordo.nome).filter(Acordo.id == id_acordo).first()                            

    return render_template ('lista_programas_acordo.html',lista_programas=lista_programas,
                                                          qtd_progs=qtd_progs,
                                                          nome = acordo.nome,
                                                          id_acordo=id_acordo)                            

#
### ATUALIZAR programa do CNPq

@acordos.route("/<int:id>/atualiza_programa_cnpq", methods=['GET', 'POST'])
@login_required
def atualiza_programa_cnpq(id):
    """
    +---------------------------------------------------------------------------------------+
    |Permite atualizar os dados de um programa do CNPq selecionado na tela de consulta.     |
    |                                                                                       |
    |Recebe o ID do programa como parâmetro.                                                |
    +---------------------------------------------------------------------------------------+
    """

    programa_cnpq = Programa_CNPq.query.get_or_404(id)

    unidade = current_user.coord

    # se unidade for pai, junta ela com seus filhos
    hierarquia = db.session.query(Coords.sigla).filter(Coords.pai == unidade).all()

    if hierarquia:
        l_unid = [f.sigla for f in hierarquia]
        l_unid.append(unidade)
    else:
        l_unid = [unidade]

    lista_coords = [(c,c) for c in l_unid]
    lista_coords.insert(0,('',''))

    form = Programa_CNPqForm()
    form.coord.choices = lista_coords

    if form.validate_on_submit():

        # atualiza acordos caso o cod. do program seja alterado
        if form.cod_programa.data != programa_cnpq.COD_PROGRAMA:
            acordos = db.session.query(Acordo).filter(Acordo.programa_cnpq==programa_cnpq.COD_PROGRAMA).all()
            if acordos:
                for a in acordos:
                    a.programa_cnpq = form.cod_programa.data
                db.session.commit()    

        programa_cnpq.COD_PROGRAMA   = form.cod_programa.data
        programa_cnpq.NOME_PROGRAMA  = form.nome_programa.data
        programa_cnpq.SIGLA_PROGRAMA = form.sigla_programa.data
        programa_cnpq.COORD          = form.coord.data

        db.session.commit()
               
        registra_log_auto(current_user.id,None,'pac')

        flash('Programa do CNPq atualizado!','sucesso')
        return redirect(url_for('acordos.lista_programa_cnpq'))
    
    # traz a informação atual do programa CNPq
    elif request.method == 'GET':

        form.cod_programa.data   = programa_cnpq.COD_PROGRAMA
        form.nome_programa.data  = programa_cnpq.NOME_PROGRAMA
        form.sigla_programa.data = programa_cnpq.SIGLA_PROGRAMA
        form.coord.data          = programa_cnpq.COORD

    return render_template('cria_programa_cnpq.html', title='Update', form=form)

#
### LISTAR processos mãe de um acordo

@acordos.route("/<int:acordo_id>/lista_processos_mae_por_acordo")
def lista_processos_mae_por_acordo(acordo_id):
    """
    +---------------------------------------------------------------------------------------+
    |Lista os processos mãe vinculados a um acordo.                                         |
    +---------------------------------------------------------------------------------------+
    """

    acordo = db.session.query(Acordo).filter(Acordo.id == acordo_id).first()

    processos = db.session.query(Acordo_ProcMae.proc_mae_id,
                                 Processo_Mae.proc_mae,
                                 Processo_Mae.coordenador,
                                 Processo_Mae.inic_mae,
                                 Processo_Mae.term_mae,
                                 Processo_Mae.situ_mae,
                                 label('qtd_filhos',func.count(Processo_Filho.processo)),
                                 label('pago',func.sum(Processo_Filho.pago_total)),
                                 label('max_ult_pag',func.max(Processo_Filho.dt_ult_pag)),
                                 Processo_Mae.pago_custeio,
                                 Processo_Mae.pago_capital,
                                 label('pago_cap_cus',Processo_Mae.pago_custeio + Processo_Mae.pago_capital),
                                 Processo_Mae.id_chamada)\
                          .join(Processo_Mae, Processo_Mae.id == Acordo_ProcMae.proc_mae_id)\
                          .outerjoin(Processo_Filho, Processo_Filho.proc_mae == Processo_Mae.proc_mae)\
                          .filter(Acordo_ProcMae.acordo_id == acordo_id)\
                          .group_by(Acordo_ProcMae.proc_mae_id,
                                    Processo_Mae.proc_mae,
                                    Processo_Mae.coordenador,
                                    Processo_Mae.inic_mae,
                                    Processo_Mae.term_mae,
                                    Processo_Mae.situ_mae,
                                    Processo_Mae.pago_custeio,
                                    Processo_Mae.pago_capital,
                                    Processo_Mae.id_chamada)\
                          .all()                      

    qtd_processos = len(processos)
    
    return render_template('lista_processos_mae.html',procs_mae=processos,
                                                      qtd_processos=qtd_processos,
                                                      chamada = None,
                                                      acordo_id=acordo.id,
                                                      acordo_tit=acordo.nome +' '+ acordo.epe +'-'+ acordo.uf)
#
### Alterar dados de um processos_mãe de um acordo

@acordos.route("/<acordo_id>/<proc_mae>/altera_mae", methods=['GET', 'POST'])
def altera_mae(acordo_id,proc_mae):
    """
    +---------------------------------------------------------------------------------------+
    |Alterar dados de um processos_mãe de um acordo.                                        |
    +---------------------------------------------------------------------------------------+
    """

    processo_mae = db.session.query(Processo_Mae).filter(Processo_Mae.proc_mae == proc_mae.replace('_','/')).first()

    form = Altera_proc_mae_Form()

    if form.validate_on_submit():

        processo_mae.coordenador = form.coordenador.data
        processo_mae.situ_mae    = form.situ_mae.data

        db.session.commit()

        registra_log_auto(current_user.id,None,'mae')

        flash('Dados de processo-mãe atualizados manualmente!','sucesso')
        return redirect(url_for('acordos.lista_processos_mae_por_acordo',acordo_id=acordo_id))

    elif request.method == 'GET':

        form.coordenador.data = processo_mae.coordenador
        form.situ_mae.data   = processo_mae.situ_mae


    return render_template('altera_mae.html', form=form, proc_mae = proc_mae)

### ASSOCIAR UM processo mãe a um acordo

@acordos.route("/<int:acordo_id>/processo_mae_acordo", methods=['GET', 'POST'])
@login_required
def processo_mae_acordo(acordo_id):
    """
    +---------------------------------------------------------------------------------------+
    |                                                                                       |
    | Permite associar um processos mãe a um acordo.                                        |
    |                                                                                       |
    | Recebe o ID do acodo como parâmetro.                                                  |
    +---------------------------------------------------------------------------------------+
    """

    # chamadas do acordo
    chamadas = db.session.query(chamadas_cnpq)\
                         .join(chamadas_cnpq_acordos, chamadas_cnpq_acordos.chamada_cnpq_id == chamadas_cnpq.id)\
                         .filter(chamadas_cnpq_acordos.acordo_id == acordo_id)\
                         .all()

    lista_chamadas = [c.id for c in chamadas]

    # processos mãe das chamadas
    maes = db.session.query(Processo_Mae).filter(Processo_Mae.id_chamada.in_(lista_chamadas))

    lista_maes = [(str(m.id),m.proc_mae) for m in maes]

    form = EscolheMaeForm()

    form.mae.choices = lista_maes

    if form.validate_on_submit():

        for mae in form.mae.data:
            acordo_procmae = Acordo_ProcMae(acordo_id   = acordo_id,
                                            proc_mae_id = int(mae))
            db.session.add(acordo_procmae)

        db.session.commit()

        registra_log_auto(current_user.id,None,'ass')

        flash('Processo Mãe relacionado ao Acordo!','sucesso')
        return redirect(url_for('acordos.lista_processos_mae_por_acordo',acordo_id=acordo_id))

    return render_template('associa_processo_mae_acordo.html', form=form,
                                                               acordo_id=acordo_id)

## registrar manualmente um processo-mãe no sistema
@acordos.route("/<int:acordo_id>/inclui_proc_mae", methods=['GET', 'POST'])
def inclui_proc_mae(acordo_id):
    """
    +---------------------------------------------------------------------------------------+
    |Registra um processos_mãe no sistema e o associa a um acordo.                          |
    +---------------------------------------------------------------------------------------+
    """

    form = Inclui_proc_mae_Form()

    if form.validate_on_submit():

        proc_mae_manual = Processo_Mae(cod_programa = None,
                                       nome_chamada = None,
                                       proc_mae     = form.proc_mae.data,
                                       inic_mae     = form.inic_mae.data,
                                       term_mae     = form.term_mae.data,
                                       coordenador  = form.coordenador.data,
                                       situ_mae     = form.situ_mae.data,
                                       id_chamada   = None,
                                       pago_capital = 0,
                                       pago_custeio = 0,
                                       pago_bolsas  = 0)                    

        db.session.add(proc_mae_manual)              
        db.session.commit()

        registra_log_auto(current_user.id,None,'mae')

        acordo_procmae = Acordo_ProcMae(acordo_id   = acordo_id,
                                        proc_mae_id = proc_mae_manual.id)

        db.session.add(acordo_procmae)
        db.session.commit()

        registra_log_auto(current_user.id,None,'ass')

        flash('Processo-mãe inseridos manualmente e relacionado ao Acordo/TED!','sucesso')


        return redirect(url_for('acordos.lista_processos_mae_por_acordo',acordo_id=acordo_id))


    return render_template('inclui_proc_mae.html', form=form,acordo_id=acordo_id)

#
### DELETAR processo MÃE de um ACORDO

@acordos.route('/<int:processo_mae_id>/<int:acordo_id>/deleta_processo_mae',methods=['GET','POST'])
def deleta_processo_mae(processo_mae_id,acordo_id):
    """
    +---------------------------------------------------------------------------------------+
    |Deleta a associação de um processo mãe com um acordo.                                  |
    |                                                                                       |
    |Recebe o id da associação como parâmetro.                                              |
    +---------------------------------------------------------------------------------------+
    """

    assoc = db.session.query(Acordo_ProcMae)\
                      .filter(Acordo_ProcMae.acordo_id == acordo_id, 
                              Acordo_ProcMae.proc_mae_id == processo_mae_id)\
                      .first()

    db.session.delete(assoc)
    db.session.commit()

    registra_log_auto(current_user.id,None,'ass')

    flash ('Associação Processo Mãe - Acordo desfeita!','sucesso')

    return redirect(url_for('acordos.lista_processos_mae_por_acordo',acordo_id=acordo_id))


### LISTAR processos filho de um processo mãe

@acordos.route("/<proc_mae>/lista_processos_filho")
def lista_processos_filho(proc_mae):
    """
    +---------------------------------------------------------------------------------------+
    |Lista os processos filho de um determinado processo mãe.                               |
    +---------------------------------------------------------------------------------------+
    """
    lista = 'mae'

    filhos_banco = db.session.query(Processo_Filho.processo,
                              Processo_Filho.nome,
                              Processo_Filho.cpf,
                              Processo_Filho.modalidade,
                              Processo_Filho.nivel,
                              Processo_Filho.situ_filho,
                              Processo_Filho.inic_filho,
                              Processo_Filho.term_filho,
                              Processo_Filho.mens_pagas,
                              Processo_Filho.pago_total,
                              Processo_Filho.dt_ult_pag)\
                       .filter(Processo_Filho.proc_mae == proc_mae.replace('_','/'))\
                       .order_by(Processo_Filho.nome,Processo_Filho.situ_filho)\
                       .all()

    qtd_filhos = db.session.query(Processo_Filho.proc_mae,
                                  label('qtd_filhos',func.count(distinct(Processo_Filho.processo))))\
                           .filter(Processo_Filho.proc_mae == proc_mae.replace('_','/'))\
                           .group_by(Processo_Filho.proc_mae)\
                           .first()
    
    if filhos_banco:
        ult_pag = [filho.dt_ult_pag for filho in filhos_banco]
        max_ult_pag = max(ult_pag).strftime("%m/%Y")
    else:
        max_ult_pag = 0

    return render_template('lista_processos_filho.html',proc_mae=proc_mae,
                                                        filhos=filhos_banco,
                                                        qtd_filhos=qtd_filhos.qtd_filhos,
                                                        lista=lista,
                                                        max_ult_pag=max_ult_pag)

#
@acordos.route("/<proc_mae>/<edic>/<epe>/<uf>/carrega_sit_sigef", methods=['GET', 'POST'])
def carrega_sit_sigef(proc_mae,edic,epe,uf):
    """
    +---------------------------------------------------------------------------------------+
    |Carrega situações dos processos-filho obtidas de uma planilha gerada via sigef         |
    +---------------------------------------------------------------------------------------+
    """

    form =  ArquivoForm()

    if form.validate_on_submit():

        tempdirectory = tempfile.gettempdir()

        f = form.arquivo.data
        fname = secure_filename(f.filename)
        arq_sigef = os.path.join(tempdirectory, fname)
        f.save(arq_sigef)

        print ('\n')
        print ('***  ARQUIVO ***',arq_sigef)

        cargaSit(arq_sigef)

        registra_log_auto(current_user.id,None,'car')

        return redirect(url_for('acordos.lista_processos_filho', proc_mae=proc_mae))

    return render_template('grab_file.html',form=form,data_ref="sigef")

#
### LISTAR bolsistas (cpf) de um processo mãe

@acordos.route("/<proc_mae>/lista_bolsistas")
def lista_bolsistas(proc_mae):
    """
    +---------------------------------------------------------------------------------------+
    |Lista bolsistas (cpfs) de um determinado processo mãe.                                 |
    +---------------------------------------------------------------------------------------+
    """

    cpfs_banco = db.session.query(Processo_Filho.nome,
                              Processo_Filho.cpf,
                              Processo_Filho.modalidade,
                              Processo_Filho.nivel,
                              Processo_Filho.situ_filho,
                              label('min_inic_filho',func.min(Processo_Filho.inic_filho)),
                              label('max_term_filho',func.max(Processo_Filho.term_filho)),
                              label('mens_p',func.sum(Processo_Filho.mens_pagas)),
                              label('pago',func.sum(Processo_Filho.pago_total)),
                              label('mens_ap',func.sum(Processo_Filho.mens_apagar)),
                              label('apagar',func.sum(Processo_Filho.valor_apagar)))\
                       .filter(Processo_Filho.proc_mae == proc_mae.replace('_','/'))\
                       .group_by(Processo_Filho.cpf,
                                 Processo_Filho.nome,
                                 Processo_Filho.modalidade,
                                 Processo_Filho.nivel,
                                 Processo_Filho.situ_filho)\
                       .order_by(Processo_Filho.nome).all()

    qtd_cpfs = len(cpfs_banco)

    cpfs = []

    for cpf in cpfs_banco:

        cpfs.append([cpf.nome,
                       cpf.cpf,
                       cpf.modalidade,
                       cpf.nivel,
                       cpf.situ_filho,
                       cpf.min_inic_filho.strftime("%x"),
                       cpf.max_term_filho.strftime("%x"),
                       cpf.mens_p,
                       locale.currency(cpf.pago, symbol=False, grouping = True),
                       cpf.mens_ap,
                       locale.currency(cpf.apagar, symbol=False, grouping = True)])



    return render_template('lista_bolsistas.html',proc_mae=proc_mae,cpfs=cpfs,
                                                   qtd_cpfs=qtd_cpfs,
                                                   prog='',edic='',epe='',uf='')
#
### LISTAR processos filhos de um acordo

@acordos.route("/<int:acordo_id>/lista_processos_filho_por_acordo")
def lista_processos_filho_por_acordo(acordo_id):
    """
    +---------------------------------------------------------------------------------------+
    |Lista os processos filhos vinculados a um acordo.                                      |
    +---------------------------------------------------------------------------------------+
    """
    lista = 'acordo'

    acordo = db.session.get(Acordo, acordo_id)

    procs_mae = db.session.query(Processo_Mae.proc_mae)\
                          .join(Acordo_ProcMae, Processo_Mae.id == Acordo_ProcMae.proc_mae_id)\
                          .filter(Acordo_ProcMae.acordo_id == acordo_id).all()

    qtd_maes = len(procs_mae)

    l_procs_mae = [p.proc_mae for p in procs_mae]

    filhos_banco = db.session.query(Processo_Filho.processo,
                                Processo_Filho.nome,
                                Processo_Filho.cpf,
                                Processo_Filho.modalidade,
                                Processo_Filho.nivel,
                                Processo_Filho.situ_filho,
                                Processo_Filho.inic_filho,
                                Processo_Filho.term_filho,
                                Processo_Filho.mens_pagas,
                                Processo_Filho.pago_total,
                                Processo_Filho.dt_ult_pag,
                                Processo_Filho.valor_apagar,
                                Processo_Filho.mens_apagar)\
                        .filter(Processo_Filho.proc_mae.in_(l_procs_mae))\
                        .order_by(Processo_Filho.situ_filho,Processo_Filho.nome).all()

    qtd_filhos = len(filhos_banco)

    if filhos_banco:
        ult_pag = [filho.dt_ult_pag for filho in filhos_banco]
        max_ult_pag = max(ult_pag).strftime("%m/%Y")
    else:
        max_ult_pag = 0    

    return render_template('lista_processos_filho.html',filhos=filhos_banco,
                                                        qtd_maes=qtd_maes,
                                                        qtd_filhos=qtd_filhos,
                                                        lista=lista,
                                                        acordo=acordo,
                                                        max_ult_pag=max_ult_pag)
#

#
### LISTAR bolsistas (cpf) de um acordo

@acordos.route("/<int:acordo_id>/lista_bolsistas")
def lista_bolsistas_acordo(acordo_id):
    """
    +---------------------------------------------------------------------------------------+
    |Lista bolsistas (cpfs) de um acordo especifico.                                        |
    +---------------------------------------------------------------------------------------+
    """

    lista = 'acordo'

    procs_mae = db.session.query(Processo_Mae.proc_mae)\
                          .join(Acordo_ProcMae, Processo_Mae.id == Acordo_ProcMae.proc_mae_id)\
                          .filter(Acordo_ProcMae.acordo_id == acordo_id).all()

    l_procs_mae = [p.proc_mae for p in procs_mae]

    qtd_maes = len(procs_mae)

    filhos = []
    qtd_filhos = 0
    ultimo_pag = []

    cpfs_banco = db.session.query(Processo_Filho.nome,
                              Processo_Filho.cpf,
                              Processo_Filho.modalidade,
                              Processo_Filho.nivel,
                              Processo_Filho.situ_filho,
                              label('min_inic_filho',func.min(Processo_Filho.inic_filho)),
                              label('max_term_filho',func.max(Processo_Filho.term_filho)),
                              label('mens_p',func.sum(Processo_Filho.mens_pagas)),
                              label('pago',func.sum(Processo_Filho.pago_total)),
                              label('mens_ap',func.sum(Processo_Filho.mens_apagar)),
                              label('apagar',func.sum(Processo_Filho.valor_apagar)))\
                       .filter(Processo_Filho.proc_mae.in_(l_procs_mae))\
                       .group_by(Processo_Filho.cpf,
                                 Processo_Filho.nome,
                                 Processo_Filho.modalidade,
                                 Processo_Filho.nivel,
                                 Processo_Filho.situ_filho)\
                       .order_by(Processo_Filho.nome).all()

    qtd_cpfs = len(cpfs_banco)

    cpfs = []

    for cpf in cpfs_banco:

        cpfs.append([cpf.nome,
                       cpf.cpf,
                       cpf.modalidade,
                       cpf.nivel,
                       cpf.situ_filho,
                       cpf.min_inic_filho.strftime("%x"),
                       cpf.max_term_filho.strftime("%x"),
                       cpf.mens_p,
                       locale.currency(cpf.pago, symbol=False, grouping = True),
                       cpf.mens_ap,
                       locale.currency(cpf.apagar, symbol=False, grouping = True)])

    return render_template('lista_bolsistas.html',proc_mae=l_procs_mae,cpfs=cpfs,
                                                   qtd_cpfs=qtd_cpfs,
                                                   prog='',edic='',epe='',uf='')
#
## RESUMO acordos

@acordos.route('/resumo_acordos')
def resumo_acordos():
    """
    +---------------------------------------------------------------------------------------+
    |Apresenta um resumo dos acordos por programa da coordenação.                           |
    |                                                                                       |
    +---------------------------------------------------------------------------------------+
    """

    unidade = current_user.coord

    # se unidade for pai, junta ela com seus filhos
    hierarquia = db.session.query(Coords.sigla).filter(Coords.pai == unidade).all()

    if hierarquia:
        l_unid = [f.sigla for f in hierarquia]
        l_unid.append(unidade)
    else:
        l_unid = [unidade]

    maes_filhos   = db.session.query(Processo_Filho.cod_programa,
                                     label('qtd_maes',func.count(distinct(Processo_Filho.proc_mae))),
                                     label('qtd_filhos',func.count(distinct(Processo_Filho.processo))),
                                     label('qtd_cpfs',func.count(distinct(Processo_Filho.cpf))),
                                     label('pago_bolsas',func.sum(Processo_Filho.pago_total)))\
                               .group_by(Processo_Filho.cod_programa)\
                               .subquery()

    chamadas = db.session.query(chamadas_cnpq.cod_programa,
                                label('pago_chamada',func.sum(chamadas_cnpq.valor)))\
                         .group_by(chamadas_cnpq.cod_programa)\
                         .subquery()                        
    
    programas = db.session.query(Programa_CNPq.COD_PROGRAMA,
                                 Programa_CNPq.SIGLA_PROGRAMA,
                                 label('vl_cnpq',func.sum(Acordo.valor_cnpq)),
                                 label('vl_epe',func.sum(Acordo.valor_epe)),
                                 label('qtd',func.count(Acordo.id)),
                                 label('qtd_edic',func.count(distinct(Acordo.nome))),
                                 maes_filhos.c.qtd_maes,
                                 maes_filhos.c.qtd_filhos,
                                 maes_filhos.c.qtd_cpfs,
                                 maes_filhos.c.pago_bolsas,
                                 chamadas.c.pago_chamada)\
                           .join(grupo_programa_cnpq, grupo_programa_cnpq.id_programa == Programa_CNPq.ID_PROGRAMA)\
                           .join(Acordo,Acordo.id == grupo_programa_cnpq.id_acordo)\
                           .outerjoin(maes_filhos, maes_filhos.c.cod_programa == Programa_CNPq.COD_PROGRAMA)\
                           .outerjoin(chamadas, chamadas.c.cod_programa == Programa_CNPq.COD_PROGRAMA)\
                           .filter(Programa_CNPq.COORD.in_(l_unid))\
                           .group_by(Programa_CNPq.SIGLA_PROGRAMA,
                                     Programa_CNPq.COD_PROGRAMA,
                                     maes_filhos.c.qtd_maes,
                                     maes_filhos.c.qtd_filhos,
                                     maes_filhos.c.qtd_cpfs,
                                     maes_filhos.c.pago_bolsas,
                                     chamadas.c.pago_chamada)\
                           .all()

    return render_template('resumo_acordos.html',programas=programas,
                                                 unidade=unidade)

#
## RESUMO por nomes dos acordos

@acordos.route('/<cod_programa>/<sigla>/edic_programa')
def edic_programa(cod_programa,sigla):
    """
    +---------------------------------------------------------------------------------------+
    |Apresenta resumo de acordos por nome, que, anteriormente era chamado de edição.        |
    |                                                                                       |
    +---------------------------------------------------------------------------------------+
    """

    chamadas = db.session.query(Acordo.nome,
                                label('pago_chamada',func.sum(chamadas_cnpq.valor)))\
                         .join(chamadas_cnpq_acordos, chamadas_cnpq_acordos.acordo_id == Acordo.id)\
                         .join(chamadas_cnpq, chamadas_cnpq.id == chamadas_cnpq_acordos.chamada_cnpq_id)\
                         .group_by(Acordo.nome)\
                         .subquery()

    maes_filhos   = db.session.query(Acordo.nome,
                                     label('qtd_maes',func.count(distinct(Processo_Filho.proc_mae))),
                                     label('qtd_filhos',func.count(distinct(Processo_Filho.processo))),
                                     label('qtd_cpfs',func.count(distinct(Processo_Filho.cpf))),
                                     label('pago_bolsas',func.sum(Processo_Filho.pago_total)))\
                               .join(Acordo_ProcMae, Acordo_ProcMae.acordo_id == Acordo.id)\
                               .join(Processo_Mae, Processo_Mae.id == Acordo_ProcMae.proc_mae_id)\
                               .join(Processo_Filho, Processo_Filho.proc_mae == Processo_Mae.proc_mae)\
                               .group_by(Acordo.nome)\
                               .subquery()                     

    edics = db.session.query(Acordo.nome,
                             label('vl_cnpq',func.sum(Acordo.valor_cnpq)),
                             label('vl_epe',func.sum(Acordo.valor_epe)),
                             label('qtd',func.count(Acordo.id)),
                             label('qtd_edic',func.count(distinct(Acordo.nome))),
                             maes_filhos.c.qtd_maes,
                             maes_filhos.c.qtd_filhos,
                             maes_filhos.c.qtd_cpfs,
                             maes_filhos.c.pago_bolsas,
                             chamadas.c.pago_chamada)\
                      .join(grupo_programa_cnpq,grupo_programa_cnpq.id_acordo == Acordo.id)\
                      .join(Programa_CNPq, Programa_CNPq.ID_PROGRAMA == grupo_programa_cnpq.id_programa)\
                      .outerjoin(maes_filhos, maes_filhos.c.nome == Acordo.nome)\
                      .outerjoin(chamadas, chamadas.c.nome == Acordo.nome)\
                      .filter(Programa_CNPq.COD_PROGRAMA == cod_programa)\
                      .group_by(Acordo.nome,
                                maes_filhos.c.qtd_maes,
                                maes_filhos.c.qtd_filhos,
                                maes_filhos.c.qtd_cpfs,
                                maes_filhos.c.pago_bolsas,
                                chamadas.c.pago_chamada)\
                      .order_by(Acordo.nome)\
                      .all()

    return render_template('edic_programa.html',edics=edics,sigla=sigla)

#
#
## acordos  no mapa do Brasil

@acordos.route('/brasil_acordos')
def brasil_acordos():
    """
    +---------------------------------------------------------------------------------------+
    |Apresenta um mapa onde se pode verificar os acordos e encomendas por UF.               |
    +---------------------------------------------------------------------------------------+
    """

    hoje = dt.today()

    gps = {'AC':[-9.977916,-67.826068],
           'AL':[-9.649433,-35.709335],
           'AM':[-3.074759,-60.028723],
           'AP':[0.052334,-51.070093],
           'BA':[-13.008304,-38.512027],
           'CE':[-3.795849,-38.497930],
           'DF':[-15.710702,-47.911077],
           'ES':[-20.276832,-40.300442],
           'GO':[-16.680903,-49.250701],
           'MA':[-2.501711,-44.284316],
           'MG':[-19.884511,-43.915749],
           'MS':[-20.447545,-54.603542],
           'MT':[-15.566057,-56.072258],
           'PA':[-1.454934,-48.475778],
           'PB':[-7.205724,-35.921335],
           'PE':[-8.060426,-34.901544],
           'PI':[-5.096300,-42.798928],
           'PR':[-25.446918,-49.245448],
           'RJ':[-22.904571,-43.173827],
           'RN':[-5.829595,-35.212017],
           'RO':[-8.625350,-63.844920],
           'RR':[2.930872,-60.672953],
           'RS':[-30.028724,-51.228277],
           'SC':[-27.571250,-48.509038],
           'SE':[-10.909057,-37.050032],
           'SP':[-23.536390,-46.714247],
           'TO':[-10.182099,-48.331027]}

    programas = db.session.query(Programa_CNPq.SIGLA_PROGRAMA,
                                 Programa_CNPq.COD_PROGRAMA,
                                 Programa_CNPq.COORD,
                                 Programa_CNPq.ID_PROGRAMA,
                                 Acordo.uf,
                                 Acordo.epe,
                                 label('qtd_acordos',func.count(Acordo.id)))\
                          .join(grupo_programa_cnpq, grupo_programa_cnpq.id_programa == Programa_CNPq.ID_PROGRAMA)\
                          .join(Acordo, Acordo.id ==  grupo_programa_cnpq.id_acordo)\
                          .filter(Acordo.data_fim >= hoje)\
                          .order_by(Programa_CNPq.SIGLA_PROGRAMA)\
                          .group_by(Acordo.uf,
                                    Programa_CNPq.SIGLA_PROGRAMA,
                                    Programa_CNPq.COD_PROGRAMA,
                                    Programa_CNPq.COORD,
                                    Programa_CNPq.ID_PROGRAMA,
                                    Acordo.epe)\
                          .all()

    m = Map(location=[-15.7, -47.9],
            tiles='OpenStreetMap',
            control_scale = True,
            zoom_start = 2,
            min_zoom=2)

    m.fit_bounds([[-34,-74],[3,-34]])

    for uf in gps:
        qtd_na_uf = 0
        cor = 'blue'
        pop = '<b>Acordos - '+uf+':</b><br>'
        for p in programas:
            if p.uf == uf:
                qtd_na_uf += p.qtd_acordos
                tip = '<b>'+uf+'</b><br>'+str(qtd_na_uf)+' acordos vigentes'
                pop += '<p>'+p.SIGLA_PROGRAMA+' ('+str(p.qtd_acordos)+')</p>'
        if pop == '<b>Acordos - '+uf+':</b><br>':
            tip = '<b>'+uf+'</b>'
            pop = 'N&atilde;o h&aacute; Acordos celebrados com <b>'+uf+'</b>.'

        #Circulos com raios em metros
        Circle(location = [float(gps[uf][0]), float(gps[uf][1])],
               radius = 18000 + qtd_na_uf*8000,
               tooltip = tip,
               popup = Popup(pop,max_width=150),
               fill = True,
               fill_opacity = 0.2,
               weight = 1,
               color=cor).add_to(m)


    return render_template('brasil_convenios.html', m = m._repr_html_())


## acordos no quadro por uf

@acordos.route('/quadro_acordos')
def quadro_acordos():
    """
    +---------------------------------------------------------------------------------------+
    |Apresenta um quadro onde se pode verificar os acordos por UF.                          |
    +---------------------------------------------------------------------------------------+
    """

    unidade = current_user.coord

    # se unidade for pai, junta ela com seus filhos
    hierarquia = db.session.query(Coords.sigla).filter(Coords.pai == unidade).all()

    if hierarquia:
        l_unid = [f.sigla for f in hierarquia]
        l_unid.append(unidade)
    else:
        l_unid = [unidade]

    acordos = db.session.query(Acordo.uf,
                               Programa_CNPq.SIGLA_PROGRAMA,
                               label('qtd',func.count(Acordo.id)),
                               label('valor_global',func.sum(Acordo.valor_cnpq+Acordo.valor_epe)),
                               Programa_CNPq.COD_PROGRAMA)\
                        .filter(Programa_CNPq.COORD.in_(l_unid),
                                or_(Acordo.situ=='Vigente-Z',Acordo.situ=='Vigente-Esquecido'))\
                        .join(grupo_programa_cnpq, grupo_programa_cnpq.id_acordo==Acordo.id)\
                        .join(Programa_CNPq, Programa_CNPq.ID_PROGRAMA == grupo_programa_cnpq.id_programa)\
                        .order_by(Acordo.uf)\
                        .group_by(Acordo.uf,Programa_CNPq.SIGLA_PROGRAMA,Programa_CNPq.COD_PROGRAMA)\
                        .all()                                        

    ufs = [a.uf for a in acordos]
    ufs = set(ufs)
    ufs = list(ufs)
    ufs.sort()

    quantidade = len(ufs)

    programas_int = [p.SIGLA_PROGRAMA+'*'+p.COD_PROGRAMA for p in acordos]
    programas_int = set(programas_int)
    programas_int = list(programas_int)
    programas_int.sort()

    linha = []
    linhas = []
    item = []

    for uf in ufs:
        for prog in programas_int:
            flag = False
            for acordo in acordos:

                if acordo.uf == uf:
                    if acordo.SIGLA_PROGRAMA == prog.split('*')[0]:
                        linha.append(acordo)
                        flag = True
                    else:
                        item = [uf,prog.split('*')[0],0,0,prog.split('*')[1]]

            if not flag:
                linha.append(item)
                flag = False

        linhas.append(linha)
        linha=[]

    return render_template('quadro_acordos.html', quantidade=quantidade,
                            programas=programas_int,linhas=linhas)

#
### gasto mês por acordo  (DESCONTINUADO)

@acordos.route("/<int:acordo_id>/<edic>/<epe>/<uf>/gasto_mes")
def gasto_mes(acordo_id,edic,epe,uf):
    """
    +---------------------------------------------------------------------------------------+
    |Lista os gastos mensais por acordo.                                                    |
    +---------------------------------------------------------------------------------------+
    """

    gastos = []

    procs_mae = db.session.query(Acordo_ProcMae.proc_mae_id,
                                 Processo_Mae.proc_mae,
                                 Processo_Mae.inic_mae,
                                 Processo_Mae.term_mae)\
                          .join(Processo_Mae, Processo_Mae.id == Acordo_ProcMae.proc_mae_id)\
                          .filter(Acordo_ProcMae.acordo_id == acordo_id).all()

    if procs_mae:                      

        for proc in procs_mae:

            pags_proc_mae = db.session.query(label('vl_pago',func.sum(PagamentosPDCTR.valor_pago)),
                                            label('qtd',func.count(PagamentosPDCTR.valor_pago)),
                                            PagamentosPDCTR.data_pagamento)\
                                    .group_by(PagamentosPDCTR.data_pagamento)\
                                    .filter(PagamentosPDCTR.proc_mae == proc.proc_mae).all()

            #datas com dias trocados para evita problemas com rrule
            strt_dt = datetime.date(proc.inic_mae.year,proc.inic_mae.month,1)
            end_dt  = datetime.date(proc.term_mae.year,proc.term_mae.month,28)

            dates = [dt for dt in rrule(MONTHLY, dtstart=strt_dt, until=end_dt)]

            for d in dates:

                d1 = datetime.date(d.year,d.month,1)

                p = 0
                q = 0
                for pag in pags_proc_mae:
                    #if pag.data_pagamento == d1:
                    if pag.data_pagamento.year == d.year and pag.data_pagamento.month == d.month:
                        p = pag.vl_pago
                        q = pag.qtd

                gastos.append((d1,p,q))

        datas = set([x[0] for x in gastos])

        lista = []
        for d in datas:
            v = 0
            qtd = 0
            for g in gastos:
                if g[0] == d:
                    v += g[1]
                    qtd += g[2]
            lista.append((d,v,qtd))

        lista.sort(key = lambda x: x[0])

        gastos = [ (x[0].year,x[0].month,locale.currency(x[1], symbol=False, grouping = True),x[2]) for x in lista ]

        return render_template('gasto_mes.html',gastos=gastos,qtd_meses=len(gastos),
                                                    prog='',edic=edic,epe=epe,uf=uf)

    else:
        flash('Não há processos-mãe registrados!','erro')

        return render_template('gasto_mes.html',gastos=[],qtd_meses=0,
                                                    prog='',edic=edic,epe=epe,uf=uf)
#
#
# pega dados de programas no DW
@acordos.route('/programas_por_unidade_DW')
@login_required
def programas_por_unidade_DW():
    """
    +---------------------------------------------------------------------------------------+
    | Alimenta tabela programas_cnpq com dados do DW                                        |
    | Chama função consultaDW, restringindo busca à registros relacionados à unidade        |
    | do usuário.                                                                           |
    +---------------------------------------------------------------------------------------+
    """

    # pega programas da unidade do usuário
    unidade = current_user.coord

    # se unidade for pai, junta ela com seus filhos
    hierarquia = db.session.query(Coords.sigla).filter(Coords.pai == unidade).all()

    if hierarquia:
        l_unid = [f.sigla for f in hierarquia]
        l_unid.append(unidade)
    else:
        l_unid = [unidade]

    u = 0 # contador de unidades
    pn = pa = 0 # contadores de programas (novos e antigos)

    for unid in l_unid:

        u += 1
        # pega no DW dados programas de cada uma das unidades identificadas
        consulta = consultaDW(tipo = 'programas_unid', unid = unid)

        for item in consulta:

            programa = db.session.query(Programa_CNPq).filter(Programa_CNPq.COD_PROGRAMA == item[0],Programa_CNPq.COORD == item[3]).first()

            if not programa:
                pn += 1
                novo_programa = Programa_CNPq(COD_PROGRAMA   = item[0],
                                              NOME_PROGRAMA  = item[1],
                                              SIGLA_PROGRAMA = item[2],
                                              COORD          = item[3])
                db.session.add(novo_programa)
            else:
                pa += 1
                programa.NOME_PROGRAMA  = item[1]
                programa.SIGLA_PROGRAMA = item[2]
                programa.COORD          = item[3]

    db.session.commit()

    flash('Efetuada carga de '+str(pn)+' programas novos e '+str(pa)+' programas já existentes'+\
          ' vinculadas à(s) '+str(u)+' unidade(s) identificadas','sucesso')

    return redirect(url_for('core.inicio'))


# lança tela de espera finalização de carga
@acordos.route('/<carga>/espera_carga')
@login_required
def espera_carga(carga):
    """
    +---------------------------------------------------------------------------------------+
    | Apresenta tela de espera carga enquato a carga e executada                            |
    +---------------------------------------------------------------------------------------+
    """

    return render_template('index_waiting.html',carga='/'+carga.replace('#','/'))



# def chamadas_DW():

#     # pega programas da unidade do usuário
#     unidade = current_user.coord

#     # se unidade for pai, junta ela com seus filhos
#     hierarquia = db.session.query(Coords.sigla).filter(Coords.pai == unidade).all()

#     if hierarquia:
#         l_unid = [f.sigla for f in hierarquia]
#         l_unid.append(unidade)
#     else:
#         l_unid = [unidade]

#     programas = db.session.query(Programa_CNPq.COD_PROGRAMA).filter(Programa_CNPq.COORD.in_(l_unid)).all()

#     cn = ca = pn = pa = fn = fm = 0 # contadores de chamadas e processos 

#     print ('*** Consulta ao DW e carga no banco do SICOPES: Chamadas, mães e filhos')

#     for p in programas:
#         # pega no DW dados de todas as chamadas de cada um dos programas identificados
#         chamadas_programas = consultaDW(tipo = 'chamadas_programas', cod_programa = p.COD_PROGRAMA)
#         print('** Programa: ',p.COD_PROGRAMA)
#         print(' ')
        
#         for cha_prog in chamadas_programas:

#             nome = cha_prog[0] +' - '+ cha_prog[1] +' - '+ cha_prog[2]   # junta tipo, sigla e nome para formar novo nome
#             id_chamada_DW = cha_prog[4]
#             chamada = db.session.query(chamadas_cnpq).filter(chamadas_cnpq.id_dw == id_chamada_DW).first()
#             # pegar somente chamadas cujos programas tenham vinculação com algum acordo
#             programas_acordos = db.session.query(Programa_CNPq)\
#                                           .join(grupo_programa_cnpq, grupo_programa_cnpq.id_programa == Programa_CNPq.ID_PROGRAMA)\
#                                           .filter(Programa_CNPq.COD_PROGRAMA == p.COD_PROGRAMA)\
#                                           .all()  
#             if programas_acordos:

#                 if not chamada:
#                     cn += 1
#                     nova_chamada = chamadas_cnpq(tipo          = cha_prog[0],
#                                                  sigla         = cha_prog[1],   
#                                                  nome          = cha_prog[2],
#                                                  valor         = cha_prog[6],  # VALOR, para VALOR_FOLHA, usar cha_prog[5]
#                                                  cod_programa  = cha_prog[7],
#                                                  id_dw         = cha_prog[4],
#                                                  qtd_processos = cha_prog[3]) 
#                     db.session.add(nova_chamada) 

#                     chamada_id = nova_chamada.id

#                 else:
#                     ca += 1
#                     chamada.tipo          = cha_prog[0]
#                     chamada.sigla         = cha_prog[1]
#                     chamada.nome          = cha_prog[2]
#                     chamada.cod_programa  = cha_prog[7]
#                     chamada.valor         = cha_prog[6] # VALOR, para VALOR_FOLHA, usar cha_prog[5]
#                     chamada.qtd_processos = cha_prog[3]

#                     chamada_id = chamada.id

#                 print('** Chamadas: ',cn,' novas ',ca,' atualizadas',' id: ',id_chamada_DW)  
#                 print(' ')  
                
#                 if chamada:

#                     chamada_cnpq_acordo = db.session.query(chamadas_cnpq_acordos).filter(chamadas_cnpq_acordos.chamada_cnpq_id == chamada_id).first()

#                     if chamada_cnpq_acordo: # pegar mães e filhos somente se a chamada já estiver relacinada a um acordo/TED

#                         if cha_prog[3] > 0:  # pega mães e filhos se ouver pelo menos um processo mãe na chamada

#                             # pega projetos vinculados à chamada e carrega em processos_mae (id_chamada recebe o seq_id_chamada)
#                             processos_chamadas = consultaDW(tipo = 'processos_chamadas', id_chamada = str(id_chamada_DW)) 
#                             # pegar processos filho de cada chamada
#                             filhos_chamadas = consultaDW(tipo = 'filhos_chamadas', id_chamada = str(id_chamada_DW))  

#                             print('** Pegando mães e fihos da chamada: ',chamada.nome)  
#                             print(' ')

#                             # varre todos os processos mãe de cada chamada oriunddos do DW para carga no banco do sicopes
#                             for pro_cha in processos_chamadas:
#                                 # ajusta conteúdo de situação caso seja nulo
#                                 if pro_cha[6]:
#                                     situ = pro_cha[6]
#                                 else: 
#                                     situ = ''  
#                                 if pro_cha[7]:
#                                     situ = situ + ' ' + pro_cha[7]
#                                 # formata número do processo mãe
#                                 proc_mae_formatado = str(pro_cha[2])[4:10]+'/'+str(pro_cha[2])[:4]+'-'+str(pro_cha[2])[10:]
#                                 # pega processos mãe conforme encontrado no DW, não existinto cria, caso contrário atualiza        
#                                 proc_mae = db.session.query(Processo_Mae).filter(Processo_Mae.proc_mae == proc_mae_formatado).first()
#                                 if not proc_mae:
#                                     pn += 1      
#                                     novo_proc_mae = Processo_Mae(cod_programa = str(pro_cha[0]),
#                                                                  nome_chamada = pro_cha[1],
#                                                                  proc_mae     = proc_mae_formatado,
#                                                                  inic_mae     = pro_cha[4],
#                                                                  term_mae     = pro_cha[5],
#                                                                  coordenador  = pro_cha[9],
#                                                                  situ_mae     = situ,
#                                                                  id_chamada   = chamada_id,
#                                                                  pago_capital = pro_cha[11],
#                                                                  pago_custeio = pro_cha[12],
#                                                                  pago_bolsas  = pro_cha[10])
#                                     db.session.add(novo_proc_mae)
#                                     id_proc_mae = novo_proc_mae.id
#                                 else:
#                                     pa += 1
#                                     proc_mae.cod_programa = str(pro_cha[0])
#                                     proc_mae.nome_chamada = pro_cha[1]
#                                     proc_mae.inic_mae     = pro_cha[4]
#                                     proc_mae.term_mae     = pro_cha[5]
#                                     #proc_mae.coordenador  = pro_cha[9]  # ver se é possível pegar coordenador via DW
#                                     proc_mae.situ_mae     = situ
#                                     proc_mae.id_chamada   = chamada_id
#                                     proc_mae.pago_capital = pro_cha[11]
#                                     proc_mae.pago_custeio = pro_cha[12]
#                                     proc_mae.pago_bolsas  = pro_cha[10]

#                                     id_proc_mae = proc_mae.id

#                                 # se a chamamada tiver só um mãe, já associa ele ao acordo
#                                 if cha_prog[3] == 1:
#                                     acordo_procmae = db.session.query(Acordo_ProcMae)\
#                                                                .filter(Acordo_ProcMae.acordo_id == chamada_cnpq_acordo.acordo_id,
#                                                                        Acordo_ProcMae.proc_mae_id ==  id_proc_mae)\
#                                                                .all()
#                                     if not acordo_procmae:
#                                         associa_acordo_procmae = Acordo_ProcMae(acordo_id   = chamada_cnpq_acordo.acordo_id,
#                                                                                 proc_mae_id = id_proc_mae)
#                                         db.session.add(associa_acordo_procmae)    
                                    
                        
#                                 # deleta todos os filhos do processo mãe da vez para carga limpa
#                                 procs_filho = db.session.query(Processo_Filho).filter(Processo_Filho.proc_mae == proc_mae_formatado).delete()
#                                 db.session.commit()

#                                 # para cada processo mãe, varre processos filho encontrados no DW
#                                 for fil_cha in filhos_chamadas:

#                                     if fil_cha[3] == pro_cha[2]:  # grava filho se ele for do processo mãe da vez

#                                         # ajusta conteúdo de situação caso seja nulo
#                                         if fil_cha[6]:
#                                             situ_filho = fil_cha[6]
#                                             if fil_cha[7]:
#                                                 situ_filho = situ_filho + ' ' + fil_cha[7]
#                                         elif fil_cha[8]:
#                                             situ_filho = fil_cha[8]
#                                         else: 
#                                             situ_filho = ''

#                                         # zerando valores nulos
#                                         if fil_cha[14]:
#                                             mens_pagas = fil_cha[14]
#                                         else:
#                                             mens_pagas = 0
#                                         if fil_cha[15]:
#                                             pago_total = fil_cha[15]
#                                         else:
#                                             pago_total = 0    

#                                         # formata número do processo filho
#                                         proc_filho_formatado = str(fil_cha[2])[4:10]+'/'+str(fil_cha[2])[:4]+'-'+str(fil_cha[2])[10:]

#                                         fn += 1
#                                         novo_proc_filho = Processo_Filho(cod_programa = fil_cha[0],
#                                                                          nome_chamada = None,
#                                                                          proc_mae     = str(fil_cha[3])[4:10]+'/'+str(fil_cha[3])[:4]+'-'+str(fil_cha[3])[10:],
#                                                                          processo     = proc_filho_formatado,
#                                                                          nome         = fil_cha[11],
#                                                                          cpf          = fil_cha[10],
#                                                                          modalidade   = fil_cha[12],
#                                                                          nivel        = fil_cha[13],
#                                                                          situ_filho   = situ_filho,
#                                                                          inic_filho   = fil_cha[4],
#                                                                          term_filho   = fil_cha[5],
#                                                                          mens_pagas   = mens_pagas,
#                                                                          pago_total   = pago_total,
#                                                                          valor_apagar = None,
#                                                                          mens_apagar  = None,
#                                                                          dt_ult_pag   = fil_cha[18])
#                                         db.session.add(novo_proc_filho)

                                            

#                                     elif fil_cha[3] == None: # verificando se ha nesta chamada processos sem mãe, pois além de mães com filho, a chamada pode ter processos de auxílio somente   
                                        
#                                         # ajusta conteúdo de situação caso seja nulo
#                                         if fil_cha[6]:
#                                             situ_filho = fil_cha[6]
#                                             if fil_cha[7]:
#                                                 situ_filho = situ_filho + ' ' + fil_cha[7]
#                                         elif fil_cha[8]:
#                                             situ_filho = fil_cha[8]
#                                         else: 
#                                             situ_filho = ''  

#                                         # zerando valores nulos
#                                         if fil_cha[15]:
#                                             pago_bolsas = fil_cha[15]
#                                         else:
#                                             pago_bolsas = 0 
#                                         if fil_cha[16]:
#                                             pago_capital = fil_cha[16]
#                                         else:
#                                             pago_capital = 0
#                                         if fil_cha[17]:
#                                             pago_custeio = fil_cha[17]
#                                         else:
#                                             pago_custeio = 0    
                                        
#                                         # formata número do processo filho
#                                         proc_filho_formatado = str(fil_cha[2])[4:10]+'/'+str(fil_cha[2])[:4]+'-'+str(fil_cha[2])[10:]

#                                         # verifia se o processo já existe na tabela de processos mãe, não existinto cria, caso contrário atualiza        
#                                         proc_mae = db.session.query(Processo_Mae).filter(Processo_Mae.proc_mae == proc_filho_formatado).first()
#                                         if not proc_mae:
#                                             novo_proc_mae = Processo_Mae(cod_programa = str(fil_cha[0]),
#                                                                          nome_chamada = fil_cha[1],
#                                                                          proc_mae     = proc_filho_formatado,
#                                                                          inic_mae     = fil_cha[4],
#                                                                          term_mae     = fil_cha[5],
#                                                                          coordenador  = fil_cha[10],
#                                                                          situ_mae     = situ_filho,
#                                                                          id_chamada   = chamada_id,
#                                                                          pago_capital = pago_capital,
#                                                                          pago_custeio = pago_custeio,
#                                                                          pago_bolsas  = pago_bolsas)
#                                             db.session.add(novo_proc_mae)
#                                             id_proc_mae = novo_proc_mae.id
#                                         else:
#                                             proc_mae.cod_programa = str(fil_cha[0])
#                                             proc_mae.nome_chamada = fil_cha[1]
#                                             proc_mae.inic_mae     = fil_cha[4]
#                                             proc_mae.term_mae     = fil_cha[5]
#                                             proc_mae.coordenador  = fil_cha[11]  
#                                             proc_mae.situ_mae     = situ_filho
#                                             proc_mae.id_chamada   = chamada_id
#                                             proc_mae.pago_capital = pago_capital
#                                             proc_mae.pago_custeio = pago_custeio
#                                             proc_mae.pago_bolsas  = pago_bolsas

#                                             id_proc_mae = proc_mae.id
#                                         fm += 1 

#                                         # se houver somente um filho sem mãe, já associa ele ao acordo
#                                         if len(filhos_chamadas) == 1:
#                                             acordo_procmae = db.session.query(Acordo_ProcMae)\
#                                                                .filter(Acordo_ProcMae.acordo_id == chamada_cnpq_acordo.acordo_id,
#                                                                        Acordo_ProcMae.proc_mae_id ==  id_proc_mae)\
#                                                                .all()
#                                             if not acordo_procmae:
#                                                 associa_acordo_procmae = Acordo_ProcMae(acordo_id   = chamada_cnpq_acordo.acordo_id,
#                                                                                         proc_mae_id = id_proc_mae)
#                                                 db.session.add(associa_acordo_procmae)


#                             print('** Mães: ',pn,' novos - ',pa,' antigos')    
#                             print('** Filhos: ',fn, ' novos' ,' mãe: ',proc_mae_formatado)
#                             print('** Processos de auxílio: ',fm)


#                         else: # se a chamada tiver 0 mães, tem que pegar processos sem mãe e os colocam como mãe no banco do sicopes  
#                             # pegar processos filho de cada chamada
#                             filhos_chamadas = consultaDW(tipo = 'filhos_chamadas', id_chamada = str(id_chamada_DW))    

#                             for fil_cha in filhos_chamadas:

#                                 if fil_cha[3] == None: # pegando somente os que não tem mãe

#                                     # ajusta conteúdo de situação caso seja nulo
#                                     if fil_cha[6]:
#                                         situ_filho = fil_cha[6]
#                                         if fil_cha[7]:
#                                             situ_filho = situ_filho + ' ' + fil_cha[7]
#                                     elif fil_cha[8]:
#                                         situ_filho = fil_cha[8]
#                                     else: 
#                                         situ_filho = ''

#                                     # zerando valores nulos
#                                     if fil_cha[15]:
#                                         pago_bolsas = fil_cha[15]
#                                     else:
#                                         pago_bolsas = 0 
#                                     if fil_cha[16]:
#                                         pago_capital = fil_cha[16]
#                                     else:
#                                         pago_capital = 0
#                                     if fil_cha[17]:
#                                         pago_custeio = fil_cha[17]
#                                     else:
#                                         pago_custeio = 0

                                    
#                                     # formata número do processo filho
#                                     proc_filho_formatado = str(fil_cha[2])[4:10]+'/'+str(fil_cha[2])[:4]+'-'+str(fil_cha[2])[10:]

#                                     # verifia se o processo já existe na tabela de processos mãe, não existinto cria, caso contrário atualiza        
#                                     proc_mae = db.session.query(Processo_Mae).filter(Processo_Mae.proc_mae == proc_filho_formatado).first()
#                                     if not proc_mae:
#                                         novo_proc_mae = Processo_Mae(cod_programa = str(fil_cha[0]),
#                                                                         nome_chamada = fil_cha[1],
#                                                                         proc_mae     = proc_filho_formatado,
#                                                                         inic_mae     = fil_cha[4],
#                                                                         term_mae     = fil_cha[5],
#                                                                         coordenador  = fil_cha[10],
#                                                                         situ_mae     = situ_filho,
#                                                                         id_chamada   = chamada_id,
#                                                                         pago_capital = pago_capital,
#                                                                         pago_custeio = pago_custeio,
#                                                                         pago_bolsas  = pago_bolsas)
#                                         db.session.add(novo_proc_mae)
#                                         id_proc_mae = novo_proc_mae.id
#                                     else:
#                                         proc_mae.cod_programa = str(fil_cha[0])
#                                         proc_mae.nome_chamada = fil_cha[1]
#                                         proc_mae.inic_mae     = fil_cha[4]
#                                         proc_mae.term_mae     = fil_cha[5]
#                                         proc_mae.coordenador  = fil_cha[11]  
#                                         proc_mae.situ_mae     = situ_filho
#                                         proc_mae.id_chamada   = chamada_id
#                                         proc_mae.pago_capital = pago_capital
#                                         proc_mae.pago_custeio = pago_custeio
#                                         proc_mae.pago_bolsas  = pago_bolsas

#                                         id_proc_mae = proc_mae.id
#                                     fm += 1 

#                                     # se houver somente um processo, já associa ele ao acordo
#                                     if len(filhos_chamadas) == 1:
#                                         acordo_procmae = db.session.query(Acordo_ProcMae)\
#                                                                .filter(Acordo_ProcMae.acordo_id == chamada_cnpq_acordo.acordo_id,
#                                                                        Acordo_ProcMae.proc_mae_id ==  id_proc_mae)\
#                                                                .all()
#                                         if not acordo_procmae:
#                                             associa_acordo_procmae = Acordo_ProcMae(acordo_id   = chamada_cnpq_acordo.acordo_id,
#                                                                                     proc_mae_id = id_proc_mae)
#                                             db.session.add(associa_acordo_procmae)

#                             print('** Processos de auxílio: ',fm)
                

#         db.session.commit()

#     return [cn,ca,pn,pa,fn,fm]

# pega dados de chamadas no DW
@acordos.route('/chamadas_por_programa_DW')
@login_required
def chamadas_por_programa_DW():
    """
    +---------------------------------------------------------------------------------------+
    | Alimenta tabela chamadas com dados do DW                                              |
    | Chama função consultaDW, restringindo busca à registros relacionados à unidade        |
    | do usuário.                                                                           |
    +---------------------------------------------------------------------------------------+
    """

    carga_chamadas = chamadas_DW()

    cn = carga_chamadas[0]
    ca = carga_chamadas[1]
    pn = carga_chamadas[2]
    pa = carga_chamadas[3]
    fn = carga_chamadas[4]
    fm = carga_chamadas[5]

    flash('Efetuada carga de '+str(cn)+' chamadas novas, '+str(pn)+' processos novos e '+str(fn)+' filhos e '+\
          'alteração de ' +str(ca)+' chamadas já existentes, '+str(pa)+' processos já existentes e ' + str(fm)+' filhos sem mãe'+\
          ' vinculadas aos programas da unidade do usuário.','sucesso')

    return redirect(url_for('core.inicio'))  
#
# pega dados financeiros de acordos no DW
@acordos.route('/dados_financeiros_acordos_DW')
@login_required
def dados_financeiros_acordos_DW():
    """
    +---------------------------------------------------------------------------------------+
    | Alimenta tabela de dados financeiros com dados do DW                                  |
    | Chama função consultaDW, restringindo busca à registros relacionados à unidade        |
    | do usuário.                                                                           |
    +---------------------------------------------------------------------------------------+
    """

    # pega programas da unidade do usuário
    unidade = current_user.coord

    # se unidade for pai, junta ela com seus filhos
    hierarquia = db.session.query(Coords.sigla).filter(Coords.pai == unidade).all()

    if hierarquia:
        l_unid = [f.sigla for f in hierarquia]
        l_unid.append(unidade)
    else:
        l_unid = [unidade]

    dfn = 0

    acordos = db.session.query(Acordo).filter(Acordo.unidade_cnpq.in_(l_unid)).all()

    for acordo in acordos:
        # pega processod de cada acordo da unidade e hierarquia do usuário logado
        processos = db.session.query(Processo_Mae.proc_mae)\
                              .join(Acordo_ProcMae, Acordo_ProcMae.proc_mae_id == Processo_Mae.id)\
                              .filter(Acordo_ProcMae.acordo_id == acordo.id)\
                              .all()
        if processos:
            l_processos_mae = [(p.proc_mae[7:11]+p.proc_mae[:6]+p.proc_mae[-1]) for p in processos]
        else:
            l_processos_mae = []                      

        processos_filho = db.session.query(Processo_Filho.processo)\
                              .join(Processo_Mae, Processo_Mae.proc_mae == Processo_Filho.proc_mae)\
                              .join(Acordo_ProcMae, Acordo_ProcMae.proc_mae_id == Processo_Mae.id)\
                              .filter(Acordo_ProcMae.acordo_id == acordo.id)\
                              .all()                
        if processos_filho:
            l_processos_filho = [(p.processo[7:11]+p.processo[:6]+p.processo[-1]) for p in processos_filho]
        else:
            l_processos_filho = []    

        l_processos = l_processos_mae + l_processos_filho

        if len(l_processos) >= 1:    

            if len(l_processos) == 1:
                l_processos = f"('{l_processos[0]}')"
            else:
                l_processos = tuple(l_processos)
            
            # consulta o DW para pegar dados financeiros da lisa de processos acima
            dados_financeiros = consultaDW(tipo = 'financeiro_processos', lista_processos = l_processos)

            financeiro = db.session.query(financeiro_acordo).filter(financeiro_acordo.id_acordo == acordo.id).delete()
            db.session.commit()
            
            for dados in dados_financeiros:
    
                novo_financeiro = financeiro_acordo(id_acordo                = acordo.id,
                                                    qtd_item_despesa         = dados[0],
                                                    valor_total_item_despesa = dados[1],
                                                    cod_fonte_recurso        = dados[2],
                                                    nme_fonte_recurso        = dados[3],
                                                    cod_plano_interno        = dados[4],
                                                    nme_plano_interno        = dados[5],
                                                    nme_categ_economica      = dados[6],
                                                    nme_natur_desp           = dados[7])
                db.session.add(novo_financeiro)
                dfn += 1

    db.session.commit()    

    flash('Realizada carga de '+str(dfn)+' registros de dados financeiros de acordos.','sucesso')

    return redirect(url_for('core.inicio'))    

# lista dados financeiros de um acordo
@acordos.route('/<int:acordo_id>/lista_dados_financeiros_acordo')
@login_required
def lista_dados_financeiros_acordo(acordo_id):
    """
    +---------------------------------------------------------------------------------------+
    | Lista dados financeiros de um acordo.                                                 |
    | Chama função consultaDW, restringindo busca à registros relacionados à unidade        |
    | do usuário.                                                                           |
    +---------------------------------------------------------------------------------------+
    """

    acordo = db.session.query(Acordo).filter(Acordo.id == acordo_id).first()

    dados_financeiros = db.session.query(financeiro_acordo).filter(financeiro_acordo.id_acordo == acordo_id).all()

    return render_template('financeiro_acordo.html',acordo = acordo, dados_financeiros=dados_financeiros)







## uso eventual para carregar chave de relacionamento com acordos e convênios em chamadas
@acordos.route('/carregaidrel')
@login_required
def carregaidrel():

    chamadas = db.session.query(Chamadas).all()
    i = 0
    for chamada in chamadas:
        acordo = db.session.query(Acordo.id).filter(Acordo.sei==chamada.sei).first()
        if acordo:
            chamada.id_relaciona = acordo.id
            i+=1
        else:
            convenio = db.session.query(DadosSEI.nr_convenio).filter(DadosSEI.sei==chamada.sei).first()
            if convenio:
                chamada.id_relaciona = 'C' + convenio.nr_convenio
                i+=1

    db.session.commit()
    print ('*** ',i,' alterações em chamadas')

    return redirect(url_for('core.inicio'))

