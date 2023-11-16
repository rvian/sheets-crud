from ctypes import resize
import PySimpleGUI as sg
from matplotlib import image
import doc_generator as dg
import excel_db as db
import ctypes
import platform
import svgUI as sv


_corEscuro1 = "#142D43"
_corCinza1 = "#5F6368"
_corFonte = "#ffffff"
_placeholderFormValues = {
        "ID": "",
        "NOME_EMPRESARIAL": "",
        "CNPJ": "",
        "CAPITAL_SOCIAL": "",
        "ENDEREÇO": "",
        "CEP": "",
        "CIDADE": "",
        "UF": "",
        "EMAIL": "",
        "TELEFONE": "",
        "BANCO": "",
        "AGÊNCIA": "",
        "N_CONTA": "",
        "REPRESENTANTE_LEGAL": "",
        "CPF_REPRESENTANTE": "",
        "ENDEREÇO_REPRESENTANTE": "",
        "CEP_REPRESENTANTE": "",
        "CIDADE_REPRESENTANTE": "",
        "UF_REPRESENTANTE": "",
        "EMAIL_REPRESENTANTE": "",
        "FONE_REPRESENTANTE": "",
        "DATA_SOLICITACAO": "",
        "DATA_RECEBIMENTO": "",
        "DATA_ASSINATURA": "",
        "DATA_EDICAO": "",
        "DATA_CRIACAO": ""
    }


# Layout
#sg.theme('DarkTeal9')
sg.set_options(font=("SegoeUI", 9), 
               border_width = 0, 
               button_element_size = (14, 4)#, button_color=sg.TRANSPARENT_BUTTON
               
               
               # CORES
            #    button_color = sg.TRANSPARENT_BUTTON, 
            #    background_color= _corEscuro1,
            #    input_elements_background_color= _corEscuro1,
            #    text_element_background_color = _corEscuro1,
            #    element_background_color = _corEscuro1,
            #    input_text_color = _corFonte,
            #    scrollbar_color = _corEscuro1
               )

selectedRow = None
criarNovo = None


def WindowCriarCRC(empresa=None):

    titulo = ""
    keySalvarEditarOuCriar = ""

    # Controla o comportamento da tela conforme for edição ou criação
    if empresa == None:
        # Define valores como vazio para não aparecer "None" nos inpus do form
        empresa = _placeholderFormValues
        titulo = "CRIAR"
        keySalvarEditarOuCriar = "SubmitSalvarCriacao"
    else:
        keySalvarEditarOuCriar = "SubmitSalvarEdicao"
        titulo = "EDITAR"

    layoutForm = [
        # placeholder para list
        [sg.Input(size=(64, 1), key='ID',
                  default_text=empresa["ID"], visible=False)],

        [sg.Text(titulo, font=('SegoeUIBold', 16), justification = 'c',
                 text_color='LightGrey', p=((14, 14), (14, 14)))],
        [sg.HorizontalSeparator()],

        # Input dados Empresa
        [sg.Text('Empresa', font=('SegoeUIBold', 14),  justification = 'c',
                 text_color='LightGrey',  p=((14, 14), (14, 2)))],

        [sg.Text('RAZÃO SOCIAL:', size=(15, 1)),
         sg.Input(size=(65, 1), key='NOME_EMPRESARIAL', default_text=empresa["NOME_EMPRESARIAL"])],

        [sg.Text('CNPJ:', size=(15, 1)),
            sg.Input(size=(25, 1), key='CNPJ', default_text=empresa["CNPJ"]), sg.VerticalSeparator(),
         sg.Text('CAPITAL SOCIAL:  '),
         sg.Input(size=(20, 1), key='CAPITAL_SOCIAL', default_text=empresa["CAPITAL_SOCIAL"])],

        # [sg.Text(('_'*40) + '       Localização     '+('_'*40), font=('SegoeUI',10), justification='center', text_color='LightGrey')],
        [sg.Text('ENDEREÇO:', size=(15, 1)),
            sg.Input(size=(65, 1), key='ENDEREÇO', default_text=empresa["ENDEREÇO"])],

        [sg.Text('CEP:', size=(15, 1)),
            sg.Input(size=(25, 1), key='CEP', default_text=empresa["CEP"]), sg.VerticalSeparator(),
         sg.Text('CIDADE: '),
            sg.Input(size=(17, 1), key='CIDADE',
                     default_text=empresa["CIDADE"]),   sg.VerticalSeparator(),
         sg.Text('UF:'),    
            sg.Input(size=(4, 1), key='UF', default_text=empresa["UF"])],

        # [sg.Text('Contato',font=('SegoeUI',16), text_color='LightGrey')],
        [sg.Text('E-MAIL:', size=(15, 1)),
            sg.Input(size=(25, 1), key='EMAIL', default_text=empresa["EMAIL"]),sg.VerticalSeparator(),
         sg.Text('TELEFONE: '),
            sg.Input(size=(25, 1), key='TELEFONE', default_text=empresa["TELEFONE"])],

        # [sg.Text('Financeiro', font=('SegoeUI',16))],
        [sg.Text('BANCO:', size=(15, 1)),
            sg.Input(size=(25, 1), key='BANCO', default_text=empresa["BANCO"]),sg.VerticalSeparator(),
         sg.Text('AGÊNCIA: '),
            sg.Input(size=(6, 1), key='AGÊNCIA',
                     default_text=empresa["AGÊNCIA"]),sg.VerticalSeparator(),
         sg.Text('CONTA: '),
            sg.Input(size=(10, 1), key='N_CONTA', default_text=empresa["N_CONTA"])],

        # Input dados representante
        [sg.Text('Representante Legal', font=('SegoeUI', 14), justification = 'c',
                 text_color='LightGrey', p=((14, 14), (14, 2)))],

        [sg.Text('NOME:', size=(15, 1)),
            sg.Input(size=(43, 1), key='REPRESENTANTE_LEGAL',
                     default_text=empresa["REPRESENTANTE_LEGAL"]),sg.VerticalSeparator(),
         sg.Text('CPF:'),
            sg.Input(size=(13, 1), key='CPF_REPRESENTANTE', default_text=empresa["CPF_REPRESENTANTE"])],

        [sg.Text('E-MAIL:', size=(15, 1)),
            sg.Input(size=(43, 1), key='EMAIL_REPRESENTANTE',
                     default_text=empresa["EMAIL_REPRESENTANTE"]),sg.VerticalSeparator(),
         sg.Text('FONE:', pad=(0, 0)),
            sg.Input(size=(13, 1), key='FONE_REPRESENTANTE', default_text=empresa["FONE_REPRESENTANTE"])],

        # [sg.Text('Localização', font=('SegoeUI',16))],
        [sg.Text('ENDEREÇO:', size=(15, 1)),
            sg.Input(size=(65, 1), key='ENDEREÇO_REPRESENTANTE', default_text=empresa["ENDEREÇO_REPRESENTANTE"])],

        [sg.Text('CEP:', size=(15, 1)),
            sg.Input(size=(13, 1), key='CEP_REPRESENTANTE',
                     default_text=empresa["CEP_REPRESENTANTE"]),sg.VerticalSeparator(),
         sg.Text('CIDADE:'),
            sg.Input(size=(28, 1), key='CIDADE_REPRESENTANTE',
                     default_text=empresa["CIDADE_REPRESENTANTE"]),sg.VerticalSeparator(),
         sg.Text('UF:'),
            sg.Input(size=(5, 1), key='UF_REPRESENTANTE', default_text=empresa["UF_REPRESENTANTE"])],

        # Input datas
        [sg.Text('Datas do Documento', font=('SegoeUI', 14), justification = 'c',
                 text_color='LightGrey', p=((14, 14), (14, 2)))],
        [sg.Text('Data solicitação (''Em:'')', size=(15, 1)),
            sg.Input(size=(15, 1), key='DATA_SOLICITACAO',
                     default_text=empresa["DATA_SOLICITACAO"]),sg.VerticalSeparator(),
         # sg.CalendarButton('📅', 'DATA_SOLICITACAO', format='%d/%m/%Y', key='C_DataRecebimento'),
         sg.Text('Data recebimento (''Recebido em:'')'),
            sg.Input(size=(15, 1), key='DATA_RECEBIMENTO', default_text=empresa["DATA_RECEBIMENTO"])],
        # sg.CalendarButton('📅', 'DATA_RECEBIMENTO', format='%d/%m/%Y', key='C_DataAssinatura'),
        [sg.Text('Data assinatura ('')'),
            sg.Input(size=(15, 1), key='DATA_ASSINATURA', default_text=empresa["DATA_ASSINATURA"])],
        # sg.CalendarButton('📅', title='Escolha a data:', no_titlebar=True, close_when_date_chosen=False, target='DATA_RECEBIMENTO', begin_at_sunday_plus=1, month_names=('Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro'), day_abbreviations=('SEG', 'TER', 'QUA', 'QUI', 'SEX', 'SAB', 'DOM'), format='%d/%m/%Y')],
        # sg.CalendarButton('📅', 'DATA_ASSINATURA', format='%d/%m/%Y', key='C_DataDocumento') ],

        [sg.Input(size=(65, 1), key='DATA_DOCUMENTO',
                  default_text=empresa["DATA_EDICAO"], visible=False)],
        [sg.Input(size=(65, 1), key='DATA_CRIACAO',
                  default_text=empresa["DATA_CRIACAO"], visible=False)],

        # Botões de ação
        [sg.HorizontalSeparator()],
        [sg.Button( key=keySalvarEditarOuCriar, image_filename = sv.btn_Salvar, button_color = sg.TRANSPARENT_BUTTON,
                   size=(14, 4), pad=((14, 14), (14, 14),)),

         sg.Button( key='SubmitGerarCRC',image_filename = sv.btn_Gerar, button_color = sg.TRANSPARENT_BUTTON,
                   size=(14, 4), pad=((14, 14), (14, 14))),
         sg.Button( key='Clear',image_filename = sv.btn_Limpar, button_color = sg.TRANSPARENT_BUTTON,
                   size=(14, 4), pad=((14, 14), (14, 14))),
         sg.Button(key='btn_Voltar',image_filename = sv.btn_Voltar, button_color = sg.TRANSPARENT_BUTTON,
                   size=(14, 4), pad=((14, 14), (14, 14)))]
    ]

    # Tela
    return sg.Window("Sheets CRUD",
                     layoutForm,
                     resizable=True,
                     element_justification='c',
                     size=(800, 800),
                     no_titlebar=True,
                     finalize=True,
                     grab_anywhere=True)


def WindowConsultarEmpresas():

    data = db.listar_empresas()
    columnName = next(iter(data[:1]), None)

    layoutLista = [

        [sg.Text('Sheets CRUD', font=('SegoeUIBold', 18),
                 text_color='LightGrey', p=((14, 14), (14, 14)))],
        [sg.HorizontalSeparator()],

        # Tabela CRCs
        [sg.Table(data[1:],
                  # auto_size_columns= True,
                  vertical_scroll_only=False,
                  headings=columnName,
                  key='CRCTable',
                  max_col_width=30, 
                  enable_events=True,
                  background_color=_corEscuro1
                  )],

        # Botões Gerar CRC, Excluir, Voltar
        [sg.Button(key='btn_Novo',image_filename = sv.btn_Novo, button_color = sg.TRANSPARENT_BUTTON),
         sg.Button(key='btn_GerarCRC',image_filename = sv.btn_Gerar, button_color = sg.TRANSPARENT_BUTTON),
         sg.Button(key='btn_EditarCRC',image_filename = sv.btn_Editar, button_color = sg.TRANSPARENT_BUTTON),
         sg.Button(key='btn_ExluirCRC',image_filename = sv.btn_Excluir, button_color = sg.TRANSPARENT_BUTTON),
         sg.Button(key='btn_Sair',image_filename = sv.btn_Sair, button_color = sg.TRANSPARENT_BUTTON)]

    ]

    # Tela
    return sg.Window("Sheets CRUD - LISTAGEM",
                     layoutLista,
                     resizable=True,
                     element_justification='c',
                     no_titlebar=True,
                     size=(800, 600),
                     finalize=True)


# Limpa os valores do form
def clear_input():
    for key in values:
        if key != 'ID':
            window[key]('')
    return None

# Ativa o DPI do PC, deixando a resolução melhor.
if int(platform.release()) >= 8:
    ctypes.windll.shcore.SetProcessDpiAwareness(True)


# Inicia janelas (apenas consultar sera construida incialmente)
windowCriarCRC, windowListarCRC = None, WindowConsultarEmpresas()


# Loop de eventos
while True:
    # values = dict
    window, event, values = sg.read_all_windows()

    def atualizarListaEmpresa():
        if windowListarCRC:
            empresas = db.listar_empresas()[1:]
            windowListarCRC['CRCTable'].update(empresas)

    # Botão Finalizar programa
    if event == sg.WINDOW_CLOSED or event == 'btn_Sair':
        break

    # Botão Voltar para lista estando na criação de CRC
    if window == windowCriarCRC and event == 'btn_Voltar':
        windowCriarCRC.hide()
        windowListarCRC.un_hide()

    # Botão Novo abre criar empresa
    if event == 'btn_Novo':
        windowCriarCRC = WindowCriarCRC()
        windowListarCRC.hide()
    
    # Botão Exluir selecionado
    if event == 'btn_ExluirCRC':

        if values["CRCTable"] == []:
            sg.popup("Selecione uma empresa.")
        else:
            row = values['CRCTable'][0] + 2
            # popup de confirmação
            confirmar = sg.popup_yes_no(
                "ATENÇÃO! \nDeseja apagar permanentemente o registro selecionado?")
            if confirmar == "Yes":
                db.excluir_empresa(row)
                atualizarListaEmpresa()

    # Botão Gerar arquivo CRC
    if window == windowCriarCRC and event == 'SubmitGerarCRC':
        dg.GerarDoc(values)

    # Botão Editar CRC selecionado
    if window == windowListarCRC and event == 'btn_EditarCRC':

        if values["CRCTable"] == []:
            sg.popup("Selecione uma empresa.")
        else:
            # atribui a var global qual a posicao do index a ser editado posteriormente.
            selectedRow = values['CRCTable'][0]
            empresaSelecionada = db.buscar_empresa(selectedRow)
            windowListarCRC.hide()
            windowCriarCRC = WindowCriarCRC(empresaSelecionada)

    # Botao Salva atualizacao CRC editado
    if event == 'SubmitSalvarEdicao':
        if values["CNPJ"] == '':
            sg.popup("Não foi possível editar a empresa! \nCNPJ não inserido. ")
        else:
            db.editar_empresa(selectedRow, values)

            atualizarListaEmpresa()
            windowCriarCRC.hide()
            windowListarCRC.un_hide()

    # Botão Limpar dados do form de criação
    if window == windowCriarCRC and event == 'Clear':
        clear_input()

    # Criar novo CRC
    if event == 'SubmitSalvarCriacao':
        if values["CNPJ"] == '':
            sg.popup("Não foi possível cadastrar a empresa! \nCNPJ não inserido. ")
        else:
            db.criar_empresa(values)
            sg.popup_auto_close("Empresa criada!", auto_close_duration=1,
                                no_titlebar=True)
            atualizarListaEmpresa()
            # Limpa o form
            clear_input()

    # Gerar CRC da empresa selecionada
    if window == windowListarCRC and event == 'btn_GerarCRC':
        if values["CRCTable"] == []:
            sg.popup("Selecione uma empresa.")
        else:
            row = values['CRCTable'][0]
            empresaSelecionada = db.buscar_empresa(row)

            dg.GerarDoc(empresaSelecionada)


window.close()
