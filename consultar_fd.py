import pandas as pd
import pyautogui as pag
import time
import pyperclip as pcl

# define pause geral
pause_geral = 2   
pag.PAUSE = pause_geral

# define lista de fd

lista_cod_fd = ['60319', '61268', '60320', '60342', '60338', '60321', '60343', '60341']
# lista_cod_fd = ['60319']

# pause personalizado 

def p_i(pause):
    pag.PAUSE = pause
    
def p_f():
    pag.PAUSE = pause_geral

def abre_wms():
    
    pag.PAUSE = pause_geral
        
    time.sleep(2)
    
    #avisa que vai começar a automação
    pag.alert('A automação vai começar, aperte em OK e NÃO mexa em nada!')
    
    #minimisa
    pag.click(x=1270, y=991)
    
    # Seleciona WMS
    
    pag.doubleClick(x=78, y=892)

    # Executa WMS
    
    p_i(0.1)
    pag.press('tab', presses=2)
    pag.press('enter')
    p_f()

    # Loguin

    login = 'renan'
    senha = 12345
    
    p_i(0.1)
    time.sleep(8)
    pcl.copy(login)
    pag.hotkey('ctrl', 'v')
    pag.press('tab')
    pcl.copy(senha)
    pag.hotkey('ctrl', 'v')

    pag.press('tab')
    pag.press('enter')
    p_f()

    # Abre gerenciamento de estoque

    p_i(0.5)
    pag.press('tab', presses=3)
    pag.press('enter')

    pcl.copy('gerenciamento')
    pag.hotkey('ctrl', 'v')

    pag.press('tab')
    pag.press('enter')
    p_f()
    
def abre_consulta_estoque_val():
    
    time.sleep(2)
    
    p_i(0.1)
    # click em 'consultas'
    pag.click(x=102, y=34)

    #click em 'consulta estoque por validade'
    pag.click(x=173, y=151)
    p_f()
    
def baixa_salvar_planilha_estoque_fd():
    
    time.sleep(2)
    lista_cod_fd = ['60319', '61268', '60320', '60342', '60338', '60321', '60343', '60341']
    caminho_pasta = r'C:\Users\estoque\Documents\estoque.renan\automação\fd_garagem_py'

    p_i(0.5)

    for i in lista_cod_fd:

        cod_fd = i

        time.sleep(2)

        #cola cod produto e baixa planilha *consulta estoque por validade*

        pcl.copy(cod_fd)
        pag.hotkey('ctrl', 'v')

        pag.press('tab', presses=4)
        pag.press('enter')

        pag.press('tab', presses=4)
        pag.press('enter')


        #salvar planilha

        p_i(0.7)

        #salvar planilha

        time.sleep(12)

        # click planilha

        pag.click(x=776, y=991)

        # fechar aviso

        time.sleep(2)
        pag.press('enter')

        # tela cheia
        pag.hotkey('ctrl', 'shift', 'f1')

        # vai em salvar como
        pag.press('alt')
        pag.press('a')
        pag.press('s')
        pag.press('tab')
        pag.press('down', presses=4)
        pag.press('enter')

        # coloca o nome do arquivo
        time.sleep(1)
        pcl.copy(i)
        pag.hotkey('ctrl', 'v')

        # vai até o caminho
        pag.press('tab', presses=11)
        pag.press('enter')


        # #copia e cola caminho
        pcl.copy(caminho_pasta)
        pag.hotkey('ctrl', 'v')
        pag.press('enter')

        # #salvar

        pag.keyDown('shift')
        pag.press('tab', presses=5)
        pag.keyUp('shift')

        pag.press('enter')

        # # #confirmar subistituição
        pag.press('tab')
        pag.press('enter')

        # #fechar
        time.sleep(2)
        pag.hotkey('alt', 'f4')
        pag.press('tab')

    p_f()

        
def abre_e_mail():
    
    destinatarios = [r'controle.estoque@grupomaxifarma.com',
                    r'recebimento@grupomaxifarma.com',
                    r'olivio.expedicao@grupomaxifarma.com']
    
    #partindo do principio que a área de trabalho esta aberta:
    #abre chrome:
    
    time.sleep(2)
    
    #minimisa
    pag.click(x=1270, y=991)
    
    #abre o chrome e abre uma nova guia
    pag.doubleClick(x=188, y=883)
    pag.hotkey('ctrl', 't')
    pag.press('f11')
    
    #abre o e-mail
    pag.click(x=44, y=18)
    time.sleep(2)
    
    #redefinir zoom
    #pag.click(x=1088, y=48)
    #pag.click(x=1110, y=86)
    
    #escrever e-mail
    
    pag.click(x=591, y=976)
    
    #cola os destinatarios
    
    for i in destinatarios:
        pag.PAUSE = 0.1
        pcl.copy(i)
        pag.hotkey('ctrl', 'v')
        pag.press('tab')
    
    #assunto
    pag.press('tab')
    pcl.copy('Relação das fraldas armazenadas na garagem')
    pag.hotkey('ctrl', 'v')
    
    #vai para o corpo do e-mail
    pag.press('tab')
    pag.press('tab')
    pag.press('tab')
    pag.press('tab')

def saudacao_obs():
    
    obs = 'Observação: São considerados apenas porta pallet e a rua 100.'
    saudacao = 'Bom dia! Segue a relação das fraldas armazenadas na garagem.'
    
    time.sleep(2)
    
    pag.PAUSE = 0.1
    
    pcl.copy(saudacao)
    pag.hotkey('ctrl', 'v')
    pag.press('enter')
    pag.press('enter')

    pcl.copy(obs)
    pag.hotkey('ctrl', 'v')
    pag.press('enter')
    pag.press('enter')
    
def mostra_analise_email():
      
    # define lista de fd na garagem

    for i in lista_cod_fd:

        df = pd.read_excel(f'{i}.xlsx')   
        
        # trata planilha
        df = df.drop(['Descrição', 'Marca', 'Num.Lote', 'Data Validade', 'Dias para o Venc.', 'Motivo Bloqueio'], axis = 1)
      
        # exclui linhas com endereço de cx fech e frac
        df = df[df['Tipo Ender.'] != 'Apanha Frac.']
        df = df[df['Tipo Ender.'] != 'Apanha Cx.Fech.']

        # separa endereços que começam com "1" de endereços que não começam com "1" "rua 100 e rua diferente de 100"
        endereços_1 = df[df['Endereço'].str.startswith('1')]
        endereços_nao_1 = df[df['Endereço'].str.startswith('0')]

        # soma as quantidades dos endereços que começam com "1" ou "rua 100"
        estoque_total_gar = endereços_1['Estoque'].sum()

        # soma as quantidades dos endereços que não começam com "1" ou diferente da "rua 100"
        estoque_total_arm = endereços_nao_1['Estoque'].sum()
        
        #mostra resultados
        pag.PAUSE = 0.1
       
        pcl.copy(f'Estoque da fralda {i} na:')
        pag.hotkey('ctrl', 'v')
        pag.press('enter')
        
        pcl.copy(f'Armazenagem = {estoque_total_arm} fardos;')
        pag.hotkey('ctrl', 'v')
        pag.press('enter')
        
        pcl.copy(f'Garagem = {estoque_total_gar} fardos.')
        pag.hotkey('ctrl', 'v')
        pag.press('enter')
        pag.press('enter')
        
    #enviar o e-mail
    pag.click(x=919, y=26)
    
abre_wms()

abre_consulta_estoque_val()
baixa_salvar_planilha_estoque_fd()

abre_e_mail()
saudacao_obs()
mostra_analise_email()