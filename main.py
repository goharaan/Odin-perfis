from selenium import webdriver
from selenium.webdriver.support.select import Select
import pandas as pd
import time as tm
import PySimpleGUI as sg
from selenium.webdriver.chrome.service import Service
import logging

def exibe_janela_login():
    sg.theme('DarkAmber')
    layout = [
                [sg.Text('Ambiente'), sg.Combo(['ODIN3 Oficial','ODIN3 Homologação'],size=(30,1))],
                [sg.Text('Usuário'), sg.InputText()],
                [sg.Text('Senha'), sg. InputText(password_char='*')],
                [sg.Button('Ok'), sg.Button('Cancel')]
             ]
    window = sg.Window('Automação de Criação de Autorizações', layout)
    eventos, valores = window.read()
    window.close()
    return eventos, valores

def exibe_janela_sistema(sistemas):
    sg.theme('DarkAmber')
    layout = [
                [sg.Text('Sistema'), sg.Combo(sistemas,size=(30,1))],
                [sg.Button('Ok'), sg.Button('Cancel')]
              ]
    window = sg.Window('Seleção de Sistema', layout)
    eventos, valores = window.read()
    window.close()
    return eventos, valores

def exibe_janela_publicacao(publicacoes):
    sg.theme('DarkAmber')
    layout = [
                [sg.Text('Publicação'), sg.Combo(publicacoes,size=(130,1))],
                [sg.Button('Ok'), sg.Button('Cancel')]
              ]
    window = sg.Window('Seleção de Publicação', layout)
    eventos, valores = window.read()
    window.close()
    return eventos, valores

def exibe_janela_perfil_data(perfis):
    sg.theme('DarkAmber')
    layout = [
                [sg.Text('Perfil'), sg.Combo(perfis,size=(30,1))],
                [sg.Text('Data de Término'), sg.Input(key='-IN5-'), sg.CalendarButton('CAL')],
                [sg.Button('Ok'), sg.Button('Cancel')]
              ]
    window = sg.Window('Seleção de Perfil e Data', layout)
    eventos, valores = window.read()
    window.close()
    return eventos, valores

def autoriza_excel(navegador, sistema, publicacao, perfil, data_termino):
    # carregando arquivo excel
    tabela = pd.read_excel("zonas.xlsx")
    for i, zona in enumerate(tabela["ZONA"]):
        zona3 = tabela.loc[i, "ZONA3"]
        zona4 = tabela.loc[i, "ZONA4"]

        # clicar botão incluir
        navegador.find_element('xpath', '//*[@id="j_idt183"]/a').click()

        # selecionando Sistema
        select_element = navegador.find_element('xpath', '//*[@id="filtroSistema"]')
        select_object = Select(select_element)
        select_object.select_by_visible_text(sistema)

        # selecionando Publicação
        select_element = navegador.find_element('xpath', '//*[@id="publicacoes"]')
        select_object = Select(select_element)
        select_object.select_by_visible_text(publicacao)

        # clique botão próximo
        tm.sleep(2)
        navegador.find_element('xpath', '//*[@id="formConsulta"]/div/div/div/div[3]/a[2]').click()
        # clique lupa
        navegador.find_element('xpath', '//*[@id="formConfig"]/div/div/div/div[2]/div[2]/div/span/a/i').click()

        # preencher campo ambiente
        navegador.find_element('xpath',
                               '//*[@id="j_idt141:j_idt143"]/div[1]/div/div/div[1]/div/div[1]/div/input').send_keys(
            str(zona4))  # colocar ZE com 4 dígitos
        # selecionando UF
        select_element = navegador.find_element('xpath',
                                                '//*[@id="j_idt141:j_idt143"]/div[1]/div/div/div[1]/div/div[2]/select')
        select_object = Select(select_element)
        select_object.select_by_value('SP')
        # selecionando Tipo
        select_element = navegador.find_element('xpath',
                                                '//*[@id="j_idt141:j_idt143"]/div[1]/div/div/div[1]/div/div[3]/select')
        select_object = Select(select_element)
        select_object.select_by_visible_text('ZONA')
        # clicar botão Pesquisar
        navegador.find_element('xpath', '//*[@id="j_idt141:j_idt143:j_idt170"]').click()
        tm.sleep(0.3)
        #VALIDAR SE O PRIMEIRO ITEM, O TEXTO CORRESPONDE A ZONA ESCOLHIDA
        zonaPesquisada = navegador.find_element('xpath', '//*[@id="j_idt141:j_idt143"]/table/tbody/tr[1]/td/a')

        while zonaPesquisada.text != zona4:
            print("ERRRROO: " + zonaPesquisada.text)
            tm.sleep(0.3)
            zonaPesquisada = navegador.find_element('xpath', '//*[@id="j_idt141:j_idt143"]/table/tbody/tr[1]/td/a')
        print(zonaPesquisada.text)
        zonaPesquisada.click()

        tm.sleep(0.3)
        # selecionando Perfil
        select_element = navegador.find_element('xpath', '//*[@id="perfilSelecionado"]')
        select_object = Select(select_element)
        select_object.select_by_visible_text(perfil)
        # Preechendo campo Data Fim
        navegador.find_element('xpath', '//*[@id="dataFim"]').send_keys(data_termino)
        # clique botão próximo
        navegador.find_element('xpath', '//*[@id="formConfig"]/div/div/div/div[3]/a[3]').click()
        # clique botão incluir
        navegador.find_element('xpath', '//*[@id="formPrincipal"]/div/div/div/div[1]/ul/li/button').click()
        # clique aba Grupo
        tm.sleep(1)
        navegador.find_element('xpath', '//*[@id="modalUsuarios"]/div/div/div[2]/ul/li[2]/a').click()
        # Preechendo campo Nome
        navegador.find_element('xpath', '//*[@id="j_idt183:frmGrupos:nmGrupo"]').send_keys(str(zona3))
        # selecionando UF
        select_element = navegador.find_element('xpath', '//*[@id="j_idt183:frmGrupos"]/div[1]/div/div/div[1]/div[2]/div[2]/select')
        select_object = Select(select_element)
        select_object.select_by_value('SP')
        # clique botão Pesquisar
        navegador.find_element('xpath', '//*[@id="j_idt183:frmGrupos:j_idt275"]').click()
        # clique no primeiro item do resultado da Pesquisa
        tm.sleep(0.3)
        #verificar se o elemento é o buscado pelo visible text
        grupoPesquisado = navegador.find_element('xpath', '//*[@id="j_idt183:frmGrupos"]/table/tbody/tr/td[1]/a')
        while grupoPesquisado.text != zona3:
            print("ERRRROO:" + grupoPesquisado.text)
            tm.sleep(0.3)
            grupoPesquisado = navegador.find_element('xpath', '//*[@id="j_idt183:frmGrupos"]/table/tbody/tr/td[1]/a')
        print(grupoPesquisado.text)
        grupoPesquisado.click()

        # clique botão próximo
        navegador.find_element('xpath', '//*[@id="formPrincipal"]/div/div/div/div[4]/a[3]').click()

        #Salvar dados no Log antes de Salvar

        ####CLIQUE BOTÃO CANCELAR
        navegador.find_element('xpath', '//*[@id="j_idt125"]/div/div/div/div[3]/a[1]').click()
        # clicar botão salvar
        #navegador.find_element('xpath','//*[@id="j_idt125"]/div/div/div/div[3]/a[3]').click()

def principal():
    # Exibe janela de Login
    event, values = exibe_janela_login()
    if event == 'Ok':
        ambiente = str(values[0])
        usuario = str(values[1])
        senha = str(values[2])

        if (ambiente == 'ODIN3 Oficial'):
            link = "https://odin3.tse.jus.br/"
        elif (ambiente == 'ODIN3 Homologação'):
            link = "https://odin3-hmg.tse.jus.br/"

        s = Service(r'./chromedriver.exe')  #abre o chromedriver pela pasta do exe
        navegador = webdriver.Chrome(service=s)
        navegador.implicitly_wait(3000)
        navegador.get(link)
        navegador.maximize_window()
        navegador.find_element('xpath', '//*[@id="username"]').send_keys(usuario)
        navegador.find_element('xpath', '//*[@id="password"]').send_keys(senha)
        navegador.find_element('xpath', '//*[@id="kc-login"]').click()

        # clicar botão Autorizações
        navegador.find_element('xpath', '//*[@id="menuItemList"]/li[7]/a').click()
        tm.sleep(1)

        # clicar botão incluir
        navegador.find_element('xpath', '//*[@id="j_idt183"]/a').click()
        tm.sleep(1)

        # Lê página e exibe janela para selecionar SISTEMA
        elem = navegador.find_element('xpath','//*[@id="filtroSistema"]')
        texto_por_linhas = elem.text.split('\n') #sem isso cada palavra fica como um item do combobox
        evento_sistema, valores_sistema = exibe_janela_sistema(texto_por_linhas)
        if evento_sistema == 'Ok':
            # selecionando Sistema
            sistema = str(valores_sistema[0]).strip()  #strip() remove espaços no início e final
            select_element = navegador.find_element('xpath', '//*[@id="filtroSistema"]')
            select_object = Select(select_element)
            select_object.select_by_visible_text(sistema)
            tm.sleep(2)

            # Lê página e exibe janela para selecionar Publicação
            elem = navegador.find_element('xpath', '//*[@id="publicacoes"]')
            texto_por_linhas = elem.text.split('\n')  # sem isso cada palavra fica como um item do combobox
            evento_publicacao, valores_publicacao = exibe_janela_publicacao(texto_por_linhas)
            if evento_publicacao == 'Ok':
                # selecionando Pulicação
                publicacao = str(valores_publicacao[0]).strip()  # strip() remove espaços no início e final
                select_element = navegador.find_element('xpath', '//*[@id="publicacoes"]')
                select_object = Select(select_element)
                select_object.select_by_visible_text(publicacao)
                tm.sleep(2)

                # clique botão próximo
                navegador.find_element('xpath', '//*[@id="formConsulta"]/div/div/div/div[3]/a[2]').click()

                #selecionando um ambiente qualquer
                # clique lupa
                navegador.find_element('xpath', '//*[@id="formConfig"]/div/div/div/div[2]/div[2]/div/span/a/i').click()
                tm.sleep(1)
                # clicar no primeiro item da tabela
                navegador.find_element('xpath', '//*[@id="j_idt141:j_idt143"]/table/tbody/tr[1]/td/a').click()

                # Ler página e exibe janela para selecionar Perfil, e Data
                elem = navegador.find_element('xpath', '//*[@id="perfilSelecionado"]')
                texto_por_linhas = elem.text.split('\n')  # sem isso cada palavra fica como um item do combobox
                evento_perfil_data, valores_perfil_data = exibe_janela_perfil_data(texto_por_linhas)
                if evento_perfil_data == 'Ok':
                    # selecionando Perfil
                    perfil = str(valores_perfil_data[0]).strip()  # strip() remove espaços no início e final
                    select_element = navegador.find_element('xpath', '//*[@id="perfilSelecionado"]')
                    select_object = Select(select_element)
                    select_object.select_by_visible_text(perfil)
                    # obtendo valor data na formatação correta
                    data_termino = str(valores_perfil_data['-IN5-'])[8:10] +'/'+ str(valores_perfil_data['-IN5-'])[5:7] +'/'+ str(valores_perfil_data['-IN5-'])[0:4]
                    tm.sleep(1)
                    # clique botão cancelar
                    navegador.find_element('xpath', '//*[@id="formConfig"]/div/div/div/div[3]/a[1]').click()

                    print('*****Iniciando leitura da planilha e criação das permissões*****')
                    print('Sistema: '+ sistema)
                    print('Publicação: '+ publicacao)
                    print('Perfil: '+ perfil)
                    print('Data de Término: ' + data_termino)
                    tm.sleep(2)

                    #Leitura do Excel e criação das autorizações
                    autoriza_excel(navegador, sistema, publicacao, perfil, data_termino)
        navegador.close()

principal()







