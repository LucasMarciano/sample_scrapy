# -*- coding: utf-8 -*-

## SCRIPT DE INTERAÇÃO COM O SITE DA GM - GPFA. - dentro da rede GM não precisa Autenticar o link

import sys  # Chamadas regulares
import os 
import re  # Expressoes regulares
import http  # HTTP Manualmente
from selenium import webdriver  # Controlador do Browser
from selenium.webdriver.common.keys import Keys  # Constantes do teclado
from urllib.parse import unquote  # Decodificação da URL
from io import BytesIO  # Trabalhar com Bytes array
from selenium.webdriver.chrome.options import Options  # Importa as opções do Chrome
import win32com.client # conecta no outlook
import time
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
import json
import pyautogui
import inspect
from paises import  mapa as mundo
from PIL import Image



current_folder = os.path.realpath(os.path.abspath(os.path.split(inspect.getfile(inspect.currentframe()))[0]))
chromedriver_path = os.path.join(current_folder, 'chromedriver.exe')
IEDriverServer_path = os.path.join(current_folder, 'IEDriverServer.exe')


#Trigger da aplicação deve ser email do outlook, antes de logar no site, deve coletar a URL do GPFA, necesário validar a token


if getattr(sys, 'frozen', False): 
     # executed as a bundled exe, the driver is in the extracted folder
        chromedriver_path = os.path.join(sys._MEIPASS, 'selenium', 'webdriver', 'chromedriver.exe')
        IEDriverServer_path = os.path.join(sys._MEIPASS, 'selenium', 'webdriver', 'IEDriverServer.exe')


# ## DICIONARIO DE PAISES, NO GPFA MOSTRA APENAS A SIGLA, 
# # #NO SITE DA FEDEX PRECISA ESTAR O NOME DO PAIS COMPLETO 
# ## SOLUÇÃO - DICIONÁRIO HASHMAP: DADOS FORAM COLETADOS DO SITE DA FEDEX


def garantia(obj):
     if obj is None:
        browser_GM.quit()  # Fecha a janela do navegador
        sys.exit(1)  # Sai do script com erro 1


def garantia_fedex(obj):
        if obj is None:
                browser.quit()  # Fecha a janela do navegador
                sys.exit(1)  # Sai do script com erro 1


#CONECTAR NO EMAIL

try:

        outlook = win32com.client.Dispatch('Outlook.Application').GetNamespace('MAPI')
        inbox = outlook.GetDefaultFolder(6).Folders.Item('GPFA')
        tratados = outlook.GetDefaultFolder(6).Folders.Item('GPFA').Folders.Item('tratados')

        messages = inbox.Items
        quantidade_de_emails = len(messages) ##FAZER O LOOPIING DO PROGRAMA ATÉ ZERAR A CAIXA

        message = messages.GetLast()  # vai coletar apenas o ultimo email
        Body_content = message.Body
        url = re.search(r'https(.*)', Body_content)
        link = url.group(0).replace('>', '')


except:
        pyautogui.alert(text='Por favor, valide se seu primeiro email na pasta GPFA é um pedido de PTA.', title='Error', button='OK')
        sys.exit()


# URL direta do GPFA com o pedido do PTA, esse link será recebido por email

opt = Options()  # objeto
browser_GM = webdriver.Ie(IEDriverServer_path)

browser_GM.get(link)

# Clicar em Details, abre o escopo do PTA

time.sleep(4)

found = browser_GM.find_element_by_css_selector('button[ng-click="viewDetail()"]').click()

time.sleep(12)

## ESTE BLOCO DE COGIDO COLETA AS INFORMAÇÕES A SEREM PREENCHIDAS NO SITE DO FEDEX
## DEPENDENDO DO Tipo De Captação OS FORMULARIOS SÃO DIFERENTES, LOGO TEM UM IF/ELSE PARA COLETAR OS DADOS

## LOCAL DE COLETA 

try:        
        validar = browser_GM.find_element_by_id('pickupPlant').is_selected()
        if validar == True:

                fornecedor =   browser_GM.find_element_by_css_selector('input[ng-model="request.model.pickup.plant.name"]')
                nome_fornecedor = fornecedor.get_attribute('value')

                endereco = browser_GM.find_element_by_css_selector('input[ng-model="request.model.pickup.plant.location.address"]')
                endereco_coleta = endereco.get_attribute('value')        

                city =  browser_GM.find_element_by_css_selector('input[ng-model="request.model.pickup.plant.location.city.name"]')
                cidade = city.get_attribute('value')

                CEP = browser_GM.find_element_by_css_selector('input[ng-model="request.model.pickup.plant.location.postalCode.name"]')
                cod_postal = CEP.get_attribute('value').replace('-', '')

                country = browser_GM.find_element_by_css_selector('input[ng-model="request.model.pickup.plant.location.country.code"]')
                pais = country.get_attribute('value')

                state = browser_GM.find_element_by_css_selector('input[ng-model="request.model.pickup.plant.location.state.name"]')
                vestado = state.get_attribute('value')

        else:
        
                fornecedor =   browser_GM.find_element_by_css_selector('input[ng-model="request.model.pickup.supplier.name"]')
                nome_fornecedor = fornecedor.get_attribute('value')

                endereco = browser_GM.find_element_by_css_selector('input[ng-model="request.model.pickup.supplier.location.address"]')
                endereco_coleta = endereco.get_attribute('value')

                city = browser_GM.find_element_by_css_selector('input[ng-model="request.model.pickup.supplier.location.city.name"]')
                cidade = city.get_attribute('value')      
                
                CEP =   browser_GM.find_element_by_css_selector('input[ng-model="request.model.pickup.supplier.location.postalCode.name"]')
                cod_postal = CEP.get_attribute('value').replace('-', '')

                country =   browser_GM.find_element_by_css_selector('input[ng-model="request.model.pickup.supplier.location.country.code"]')
                pais = country.get_attribute('value')
                
                state = browser_GM.find_element_by_css_selector('input[ng-model="request.model.pickup.supplier.location.state.name"]')
                vestado = state.get_attribute('value')




        contato =   browser_GM.find_element_by_css_selector('input[ng-model="request.model.pickup.contact.name"]')
        nome_contato = contato.get_attribute('value')
        tel = browser_GM.find_element_by_css_selector('input[ng-model="request.model.pickup.contact.phone"]')
        telefone = tel.get_attribute('value').replace(' ', '').replace('-', '').replace('+', '')
        PTA = browser_GM.find_element_by_class_name('detailSubHeader').text


        try:
                if len(vestado) > 3:
                        estado = vestado 
                else:
                        estado = mundo[pais]['states'][vestado]

                pais_coleta = mundo[pais]['name']

        except:
                pyautogui.alert(text='Estado não localizado, reveja o pedido de PTA.', title='Error', button='OK')
                browser_GM.quit()
                sys.exit()

       

except: # refaz a primeira etapa caso o GPFA abra o formulario em branco

        browser_GM.quit()
        
        outlook = win32com.client.Dispatch('Outlook.Application').GetNamespace('MAPI')
        inbox = outlook.GetDefaultFolder(6).Folders.Item('GPFA') 

        messages = inbox.Items
        message = messages.GetLast()  # vai coletar apenas o ultimo email
        Body_content = message.Body

        url = re.search(r'https(.*)', Body_content)
        link = url.group(0).replace('>', '')

        opt = Options()  # objeto
        browser_GM = webdriver.Ie(IEDriverServer_path)

        browser_GM.get(link)
        time.sleep(3)
        found =  browser_GM.find_element_by_css_selector('button[ng-click="viewDetail()"]').click()
        time.sleep(12)


        validar = browser_GM.find_element_by_id('pickupPlant').is_selected()
        if validar == True:

                fornecedor =   browser_GM.find_element_by_css_selector('input[ng-model="request.model.pickup.plant.name"]')
                nome_fornecedor = fornecedor.get_attribute('value')

                endereco = browser_GM.find_element_by_css_selector('input[ng-model="request.model.pickup.plant.location.address"]')
                endereco_coleta = endereco.get_attribute('value')        

                city =  browser_GM.find_element_by_css_selector('input[ng-model="request.model.pickup.plant.location.city.name"]')
                cidade = city.get_attribute('value')

                CEP = browser_GM.find_element_by_css_selector('input[ng-model="request.model.pickup.plant.location.postalCode.name"]')
                cod_postal = CEP.get_attribute('value').replace('-', '')

                country = browser_GM.find_element_by_css_selector('input[ng-model="request.model.pickup.plant.location.country.code"]')
                pais = country.get_attribute('value')

                state = browser_GM.find_element_by_css_selector('input[ng-model="request.model.pickup.plant.location.state.name"]')
                vestado = state.get_attribute('value')

        else:
        
                fornecedor =   browser_GM.find_element_by_css_selector('input[ng-model="request.model.pickup.supplier.name"]')
                nome_fornecedor = fornecedor.get_attribute('value')

                endereco = browser_GM.find_element_by_css_selector('input[ng-model="request.model.pickup.supplier.location.address"]')
                endereco_coleta = endereco.get_attribute('value')

                city = browser_GM.find_element_by_css_selector('input[ng-model="request.model.pickup.supplier.location.city.name"]')
                cidade = city.get_attribute('value')      
                
                CEP =   browser_GM.find_element_by_css_selector('input[ng-model="request.model.pickup.supplier.location.postalCode.name"]')
                cod_postal = CEP.get_attribute('value').replace('-', '')

                country =   browser_GM.find_element_by_css_selector('input[ng-model="request.model.pickup.supplier.location.country.code"]')
                pais = country.get_attribute('value')
                
                state = browser_GM.find_element_by_css_selector('input[ng-model="request.model.pickup.supplier.location.state.name"]')
                vestado = state.get_attribute('value')


contato =   browser_GM.find_element_by_css_selector('input[ng-model="request.model.pickup.contact.name"]')
nome_contato = contato.get_attribute('value')
tel = browser_GM.find_element_by_css_selector('input[ng-model="request.model.pickup.contact.phone"]')
telefone = tel.get_attribute('value').replace(' ', '').replace('-', '').replace('+', '')
PTA = browser_GM.find_element_by_class_name('detailSubHeader').text


try:
        if len(vestado) > 3:
                estado = vestado 
        else:
                estado = mundo[pais]['states'][vestado]

        pais_coleta = mundo[pais]['name']

except:
        pyautogui.alert(text='Estado não localizado, reveja o pedido de PTA.', title='Error', button='OK')
        browser_GM.quit()
        sys.exit()



## LOCAL DE ENTREGA 
## DEPENDENDO DO TIPO CAPITAÇÃO, O FORMULARIO É DIFERENTE


validar = browser_GM.find_element_by_id('request.model.deliveryPlant').is_selected()
if validar == True:

        predio =  browser_GM.find_element_by_css_selector('input[ng-model="request.model.delivery.plant.name"]')
        planta = predio.get_attribute('value')

        ende_planta =  browser_GM.find_element_by_css_selector('input[ng-model="request.model.delivery.plant.location.address"]' )
        endereco_planta = ende_planta.get_attribute('value')

        city_planta =   browser_GM.find_element_by_css_selector('input[ng-model="request.model.delivery.plant.location.city.name"]')
        cidade_planta = city_planta.get_attribute('value')

        CEP_entrega =  browser_GM.find_element_by_css_selector('input[ng-model="request.model.delivery.plant.location.postalCode.name"]')
        cod_postal_entrega = CEP_entrega.get_attribute('value').replace('-', '')

        country_entrega = browser_GM.find_element_by_css_selector('input[ng-model="request.model.delivery.plant.location.country.code"]')
        pais_entrega = country_entrega.get_attribute('value')
        pais_coleta_sigla = mundo[pais_entrega]['name']


        state_entrega =   browser_GM.find_element_by_css_selector('input[ng-model="request.model.delivery.plant.location.state.name"]')
        vestado_entrega = state_entrega.get_attribute('value')


        if len(vestado_entrega) > 3:
                estado_entrega = vestado_entrega 
        else:
                estado_entrega = mundo[pais_entrega]['states'][vestado_entrega]


        requisitante =  browser_GM.find_element_by_css_selector('input[ng-model="request.model.contact.name"]')
        nome_requisitante = requisitante.get_attribute('value')

        tel_requisitante =  browser_GM.find_element_by_css_selector('input[ng-model="request.model.contact.phone"]')
        telefone_requisitante = tel_requisitante.get_attribute('value').replace(' ', '').replace('-', '').replace('+', '')


else:
        predio =  browser_GM.find_element_by_css_selector('input[ng-model="request.model.delivery.supplier.name"]')
        planta = predio.get_attribute('value')

        ende_planta =  browser_GM.find_element_by_css_selector('input[ng-model="request.model.delivery.supplier.location.address"]' )
        endereco_planta = ende_planta.get_attribute('value')

        city_planta =   browser_GM.find_element_by_css_selector('input[ng-model="request.model.delivery.supplier.location.city.name"]')
        cidade_planta = city_planta.get_attribute('value')


        CEP_entrega =  browser_GM.find_element_by_css_selector('input[ng-model="request.model.delivery.supplier.location.postalCode.name"]')
        cod_postal_entrega = CEP_entrega.get_attribute('value').replace('-', '')

        country_entrega = browser_GM.find_element_by_css_selector('input[ng-model="request.model.delivery.supplier.location.country.code"]')
        pais_entrega = country_entrega.get_attribute('value')
        pais_coleta_sigla = mundo[pais_entrega]['name']


        state_entrega =   browser_GM.find_element_by_css_selector('input[ng-model="request.model.delivery.supplier.location.state.name"]')
        vestado_entrega = state_entrega.get_attribute('value')


        if len(vestado_entrega) > 3:
                estado_entrega = vestado_entrega 
        else:
                estado_entrega = mundo[pais_entrega]['states'][vestado_entrega]



        requisitante =  browser_GM.find_element_by_css_selector('input[ng-model="request.model.contact.name"]')
        nome_requisitante = requisitante.get_attribute('value')

        tel_requisitante =  browser_GM.find_element_by_css_selector('input[ng-model="request.model.contact.phone"]')
        telefone_requisitante = tel_requisitante.get_attribute('value').replace(' ', '').replace('-', '').replace('+', '')




                                      # ######### DETALHES DO EMBARQUE ##########


vpeso = browser_GM.find_element_by_css_selector('input[ng-model="request.model.weight"]')
peso_total = vpeso.get_attribute('value').replace(',','.')  # #UNIDADE DE MEDIDA DE PESO SEMPRE EM KILOS

custo = browser_GM.find_element_by_css_selector('input[ng-model="request.model.estimatedCost"]')
custo_estimado = custo.get_attribute('value')


##DETALHES DO EMBARQUE (QUEM PAGA PELO ENVIO) ##

select =  Select(browser_GM.find_element_by_css_selector('select[ng-init="getFreight()"]'))
frete = select.first_selected_option.text


faturar_para = browser_GM.find_element_by_css_selector('input[ng-model="request.model.billTo.accountNumber"]')
vfaturar_para = faturar_para.get_attribute('value').replace(' ','')


                                               ########## DETALHES DA PEÇA ###########
count_part =   browser_GM.find_element_by_css_selector('input[ng-model="part.quantity"]')
quantidade_peca = count_part.get_attribute('value')


descricao_peca =  browser_GM.find_element_by_css_selector('input[ng-model="part.description"]')
peca = descricao_peca.get_attribute('value')


preco_peca =  browser_GM.find_element_by_css_selector('input[ng-model="part.carline"]')
veiculo = preco_peca.get_attribute('value')


tipo_embalagem =  browser_GM.find_element_by_css_selector('input[ng-model="part.containerType"]')
embalagem = tipo_embalagem.get_attribute('value')


count_embalagem = browser_GM.find_element_by_css_selector('input[ng-model="part.containerQuantity"]')
quantidade_por_embalagem = count_embalagem.get_attribute('value')


quantidade_embalagem =  browser_GM.find_element_by_css_selector('input[ng-model="package.quantity"]')
numero_pacotes = quantidade_embalagem.get_attribute('value')


##DIMENSAO DA CAIXA 
# padrão EUA, trocar virgula por ponto

L = browser_GM.find_element_by_css_selector('input[ng-model="package.length"]')
lenght  = L.get_attribute('value').replace(',','.')

P = browser_GM.find_element_by_css_selector('input[ng-model="package.width"]')
width = P.get_attribute('value').replace(',','.')

A =  browser_GM.find_element_by_css_selector('input[ng-model="package.height"]')
height = A.get_attribute('value').replace(',','.')


browser_GM.quit()



 #ESTE BLOCO DE CODIGO ABRE O SITE DA FEDEX
 #E COMEÇA A POPULAR OS DADOS DO GPFA

# URL direta do serviço de etiquetas Fedex

browser = webdriver.Chrome(chromedriver_path)  # Cria a janela do navegador Site Fedex
browser.get('https://www.fedex.com/fcl/;SHIPPINGSESSIONID=rL6E37Oj4ObzMGiNOgol1EHaqH4zBKMWXGLVfRxfCiVPH9OgIVXT!-280826187?appName=fclfsm&locale=br_pt&step3URL=https%3A%2F%2Fwww.fedex.com%2Fshipping%2FshipEntryAction.do%3Fmethod%3DdoRegistration%26link%3D1%26locale%3Dpt_BR%26urlparams%3Dbr%26sType%3DF&returnurl=https%3A%2F%2Fwww.fedex.com%2Fshipping%2FshipEntryAction.do%3Fmethod%3DdoEntry%26link%3D1%26locale%3Dpt_BR%26urlparams%3Dbr%26sType%3DF&programIndicator=0'
            )

##Fazer login no site da fedex

found = browser.find_element_by_css_selector('input[title="ID de Usuário"]')
garantia_fedex(found)
found.send_keys('user')
found.send_keys(Keys.TAB)


found = browser.find_element_by_css_selector('input[title="Senha"]')
garantia_fedex(found)
found.send_keys('password')
found.send_keys(Keys.ENTER)

# Bloco para clicar no editar e alterar o nome do remetente, por default vem com o nome do user

el = browser.find_elements_by_class_name('adaLinkWrapperEdit')  # mais de um elemento com a mesma classe
found = None
for e in el:
    if e.text == 'Editar':
        found = e
        break
garantia_fedex(found)
found.click()

time.sleep(3)
#  1. De   

##PAÍS - precisa usar o dicionario hashmap

select = Select(browser.find_element_by_id('fromData.countryCode'))
select.select_by_visible_text(pais_coleta)


# empresa

found = browser.find_element_by_id('fromData.companyName')
garantia_fedex(found)
found.clear()
found.click()
found.send_keys(nome_fornecedor)

##nome de contato

found = browser.find_element_by_id('fromData.contactName')
garantia_fedex(found)
found.clear()
found.click()
found.send_keys(nome_contato)

##endereço

found = browser.find_element_by_id('fromData.addressLine1')
garantia_fedex(found)
found.clear()
found.send_keys(endereco_coleta)

##CEP

found = browser.find_element_by_id('fromData.zipPostalCode')
garantia_fedex(found)
found.clear()

found.send_keys(cod_postal)

## cidade

found = browser.find_element_by_id('fromData.city')
garantia_fedex(found)
found.clear()
found.send_keys(cidade)

##telefone

found = browser.find_element_by_id('fromData.phoneNumber')
garantia_fedex(found)
found.clear()
found.send_keys(telefone)

time.sleep(1)


##telefone (bug do site, não registra o valor)
found = browser.find_element_by_id('fromData.phoneNumber')
garantia_fedex(found)
found.clear()
found.send_keys(telefone)

time.sleep(1)


##Estado
select = Select(browser.find_element_by_id('fromData.stateProvinceCode'))

try:
        select.select_by_visible_text(estado)
except:
        pass #falha silencionamente




# 2. PARA   

##Pais
select = Select(browser.find_element_by_id('toData.countryCode'))
select.select_by_visible_text(pais_coleta_sigla)


##Empresa

found = browser.find_element_by_name('toData.addressData.companyName')
garantia_fedex(found)
found.click()
found.clear()
found.send_keys(planta)

##Nome de contato

found = browser.find_element_by_id('toData.contactName')
garantia_fedex(found)
found.click()
found.clear()
found.send_keys(nome_requisitante)

##endereço

found = browser.find_element_by_id('toData.addressLine1')
garantia_fedex(found)
found.click()
found.clear()
found.send_keys(endereco_planta)

##Cep

found = browser.find_element_by_id('toData.zipPostalCode')
garantia_fedex(found)
found.click()
found.clear()
found.send_keys(cod_postal_entrega)

time.sleep(1)

##Cidade

found = browser.find_element_by_id('toData.city')
garantia_fedex(found)
found.click()
found.clear()
found.send_keys(cidade_planta)

##Municipio: de acordo com o endereço, nem sempre este campo estará visivel.

try:
        found = browser.find_element_by_id('toData.stateProvinceCode')
        garantia_fedex(found)
        found.send_keys(estado_entrega)
except:
        pass #falha silencionamente



found = browser.find_element_by_id('toData.city')
garantia_fedex(found)
found.click()
found.clear()
found.send_keys(cidade_planta)


##Telefone

found = browser.find_element_by_id('toData.phoneNumber')
garantia_fedex(found)
found.clear()
found.send_keys(telefone_requisitante)

time.sleep(1)

found = browser.find_element_by_id('toData.phoneNumber')
garantia_fedex(found)
found.clear()
found.send_keys(telefone_requisitante)

time.sleep(1)

  
                ############# 4. Detalhes de Cobrança #############

###############  VFATURAR PARA NÃO ESTA SENDO COLETADO......................

##Faturar o transporte para

## se frete = collect and vfaturar_para  for nulo, valor deve ser pago pela planta da GM SCS (my account)
## se frete = collect and vfaturar_para  tiver um valor, pagamento fica com a GM mas outra planta (terceiros)
## se frete = prepaind valor deve ser pago pelo fornecedor (destinatário)

## campo Sua referência sempre informar o numero do PTA

if frete == 'Collect' and vfaturar_para == '':
        select =  Select(browser.find_element_by_id('billingData.transportationChargesBillingTypeCode'))
        select.select_by_visible_text('My account-670')
        time.sleep(2)
        select =  Select(browser.find_element_by_id('billingData.dutiesAndTaxesBillingTypeCode'))
        select.select_by_visible_text('My account-670')
        time.sleep(2)
        found = browser.find_element_by_id('billingData.yourReference')
        garantia_fedex(found)
        found.send_keys(PTA)

elif frete == 'Collect' and vfaturar_para != '':
        select =  Select(browser.find_element_by_id('billingData.transportationChargesBillingTypeCode'))
        select.select_by_visible_text('Terceiros')
        time.sleep(2)
        # numero da conta
        found = browser.find_element_by_id('billingData.transportationChargesBillingAccountInfo.accountNumber')
        garantia_fedex(found)
        found.send_keys(vfaturar_para)
        #cobrar taxas
        select =  Select(browser.find_element_by_id('billingData.dutiesAndTaxesBillingTypeCode'))
        select.select_by_visible_text('Terceiros')
        time.sleep(2)
        
        ##Nº da conta
        found = browser.find_element_by_id('billingData.dutiesAndTaxesBillingAccountInfo.accountNumber')
        garantia_fedex(found)
        found.send_keys(vfaturar_para)
        
        ##referencia
        found = browser.find_element_by_id('billingData.yourReference')
        garantia_fedex(found)
        found.send_keys(PTA)

elif frete == 'Prepaid':
        select =  Select(browser.find_element_by_id('billingData.transportationChargesBillingTypeCode'))
        select.select_by_visible_text('Destinatário')
        time.sleep(2)
        # numero da conta
        found = browser.find_element_by_id('billingData.transportationChargesBillingAccountInfo.accountNumber')
        garantia_fedex(found)
        found.send_keys(vfaturar_para)
        #cobrar taxas
        select =  Select(browser.find_element_by_id('billingData.dutiesAndTaxesBillingTypeCode'))
        select.select_by_visible_text('Destinatário')
        time.sleep(2)
        
        ##Nº da conta
        found = browser.find_element_by_id('billingData.dutiesAndTaxesBillingAccountInfo.accountNumber')
        garantia_fedex(found)
        found.send_keys(vfaturar_para)
        
        ##referencia
        found = browser.find_element_by_id('billingData.yourReference')
        garantia_fedex(found)
        found.send_keys(PTA)

    
##########5. Coleta/Entrega

time.sleep(1)
found = browser.find_element_by_id('pdm.initialChoice.useScheduledPickup').click()

##########   3. Detalhes do Pacote e da Remessa  ########## 

# documento náo tem quantidade no PTA
#Tipo serviço - se as dimensões da caixa ultrapassar 300 polegadas, 
# o tipo de serviço deve ser International Priority Freight else International Priority
#Número de confirmação da reserva 4807432412 se as dimensões da caixa ultrapassar 300 polegadas
# ou é peça ou documento, se doc preecher as infos e clicar em continuar, não troca de tela 
# se peça, troca de tela para preencher a peça e depois finaliza
# bloco de cod abaixo faz essa validação e preenchimento


if  re.search('doc', peca) != None: ##valida se é peça ou docummento        
        select = Select(browser.find_element_by_id('psdData.numberOfPackages'))
        select.select_by_visible_text('1') #numero de pacores

        found = browser.find_element_by_id('psd.mps.row.weight.0')
        garantia_fedex(found)
        found.send_keys('0.5') #documenbto tem o peso 0.5 sempre

        select = Select(browser.find_element_by_id('psdData.weightUnitOfMeasureLine1'))
        select.select_by_visible_text('kg') #unidade de medida de peso sempre deve ser em kg

        found = browser.find_element_by_id('psd.mps.row.declaredValue.0')
        garantia_fedex(found)
        found.send_keys('1') #documenbto sempre é 1 dolar

        select = Select(browser.find_element_by_id('psdData.declaredValueCurrencyCodeLine1'))
        time.sleep(2)
        select.select_by_visible_text('Dólar (EUA)') #sempre em dólar

        select = Select(browser.find_element_by_id('psdData.serviceType'))
        select.select_by_visible_text('International Priority')  # TIPO DE SERVIÇO 

        time.sleep(1)
        select = Select(browser.find_element_by_id('psdData.packageType'))
        select.select_by_visible_text('Envelope FedEx')  #doc sempre é envelope Fedex

        found = browser.find_element_by_id('commodityData.packageContents.documents').click() #conteúdo do pacote


        select = Select(browser.find_element_by_id('commodityData.documentDescriptionCode'))
        select.select_by_visible_text('Descrição do seu documento') #doc sempre usa esta opção

        found = browser.find_element_by_id('commodityData.yourDocumentDescription')
        garantia_fedex(found)
        found.send_keys('Documento')  


        found = browser.find_element_by_id('commodityData.totalCustomsValue')
        garantia_fedex(found)
        found.clear()   
        found.click()
        found.send_keys('1') 
        found.send_keys(Keys.TAB)


        time.sleep(3)
        found = browser.find_elements_by_css_selector('input[class="fsmContentFull"]')
        
        time.sleep(2)
        browser.find_element_by_xpath(".//*[@id='customsdocuments.CI.chbx']").click()
        time.sleep(2)

        browser.find_element_by_xpath(".//*[@id='customs.documents.checkbox.letterHeadImage']").click()
        browser.find_element_by_xpath(".//*[@id='customs.documents.checkbox.signatureImage']").click()
        
        found = browser.find_element_by_css_selector('input[value="Enviar"]').click() # clica em enviar [troca de tela]
        time.sleep(2) # aguarda carregar a proxima tela

        found = browser.find_element_by_css_selector('input[value="Enviar"]').click() # clica em enviar [troca de tela]
        

        found = browser.find_element_by_css_selector('input[value="Imprimir"]').click() # imprime todas as pags do PTA
        time.sleep(5)

        pyautogui.press('enter')  # Preciona tecla enter
                                  # impressora PDF deve estar setada como padrão
        time.sleep(3)
        pyautogui.press('enter')
        pyautogui.hotkey('tab')
        time.sleep(1)
        pyautogui.hotkey('shift','tab')
        
        time.sleep(1)
        pyautogui.typewrite("C:\\Users\\Public\\FEDEX - GM\\"+PTA+'.pdf')
        time.sleep(1)
        pyautogui.press('enter')
        
        time.sleep(10)
        
        found = browser.find_element_by_css_selector('img[alt="labelImage"]') #download da imagem

        time.sleep(2)


        try:
                found.screenshot(PTA + '.png')
        except Exception as ex:
                browser.get_screenshot_as_file(PTA + '.png')

        rgba = Image.open(PTA + '.png')
        rgb = Image.new('RGB', rgba.size, (255, 255, 255))  
        rgb.paste(rgba, mask=rgba.split()[3])               
        rgb.save(PTA + '.pdf', 'PDF', resoultion=100.0)
        os.remove(PTA+'.png')

else:
        if float(lenght) + float(width) + float(height) < 300:
                select = Select(browser.find_element_by_id('psdData.numberOfPackages'))
                select.select_by_visible_text(numero_pacotes) #numero de pacores

                found = browser.find_element_by_id('psd.mps.row.weight.0')
                garantia_fedex(found)
                found.send_keys(peso_total) #documenbto tem o peso 0.5 sempre

                select = Select(browser.find_element_by_id('psdData.weightUnitOfMeasureLine1'))
                select.select_by_visible_text('kg') #unidade de medida de peso sempre deve ser em kg

                found = browser.find_element_by_id('psd.mps.row.declaredValue.0')
                garantia_fedex(found)
                found.send_keys(veiculo) # custo

                select = Select(browser.find_element_by_id('psdData.declaredValueCurrencyCodeLine1'))
                select.select_by_visible_text('Dólar (EUA)') #sempre em dólar
   
                select = Select(browser.find_element_by_id('psdData.serviceType'))
                select.select_by_visible_text('International Priority' )  # TIPO DE SERVIÇO
                
                if embalagem == 'Box':
                        select = Select(browser.find_element_by_id('psdData.packageType'))
                        select.select_by_visible_text('FedEx Box')
                else:
                        select = Select(browser.find_element_by_id('psdData.packageType'))
                        select.select_by_visible_text('Sua Embalagem')
                        select = Select(browser.find_element_by_id('psd.mps.row.dimensions.0'))
                        select.select_by_visible_text('Informar dimensões')

                        found = browser.find_element_by_id('psd.mps.row.dimensionLength.0')
                        found.send_keys(lenght)
                        found = browser.find_element_by_id('psd.mps.row.dimensionWidth.0')
                        found.send_keys(width)
                        found = browser.find_element_by_id('psd.mps.row.dimensionHeight.0')
                        found.send_keys(height)
                        found = browser.find_element_by_id('commodityData.packageContents.products').click() #produtos/mercadorias

                        found = browser.find_element_by_id('commodityData.totalCustomsValue')
                        found.send_keys(veiculo) #custo
                        select = Select(browser.find_element_by_id('commodityData.totalCustomsValueCurrencyCode'))
                        select.select_by_visible_text('Dólar (EUA)') #moeda sempre em dolar

        else:
                select = Select(browser.find_element_by_id('psdData.numberOfPackages'))
                select.select_by_visible_text(numero_pacotes) #numero de pacores

                found = browser.find_element_by_id('psd.mps.row.weight.0')
                garantia_fedex(found)
                found.send_keys(peso_total) #documenbto tem o peso 0.5 sempre

                select = Select(browser.find_element_by_id('psdData.weightUnitOfMeasureLine1'))
                select.select_by_visible_text('kg') #unidade de medida de peso sempre deve ser em kg

                found = browser.find_element_by_id('psd.mps.row.declaredValue.0')
                garantia_fedex(found)
                found.send_keys(veiculo) # custo

                select = Select(browser.find_element_by_id('psdData.declaredValueCurrencyCodeLine1'))
                select.select_by_visible_text('Dólar (EUA)') #sempre em dólar
   
                select = Select(browser.find_element_by_id('psdData.serviceType'))
                select.select_by_visible_text('International Priority Freight')  # TIPO DE SERVIÇO
                select = Select(browser.find_element_by_id('psdData.packageType'))
                select.select_by_visible_text('Sua Embalagem')
                select = Select(browser.find_element_by_id('psd.mps.row.dimensions.0'))
                select.select_by_visible_text('Informar dimensões')
                found = browser.find_element_by_id('psd.mps.row.dimensionLength.0')
                found.send_keys(lenght)
                found = browser.find_element_by_id('psd.mps.row.dimensionWidth.0')
                found.send_keys(width)
                found = browser.find_element_by_id('psd.mps.row.dimensionHeight.0')
                found.send_keys(height)
                found = browser.find_element_by_id('commodityData.packageContents.products').click() #produtos/mercadorias
                found = browser.find_element_by_id('commodityData.totalCustomsValue')
                found.send_keys(veiculo) #custo
                select = Select(browser.find_element_by_id('commodityData.totalCustomsValueCurrencyCode'))
                select.select_by_visible_text('Dólar (EUA)') #moeda sempre em dolar

        # clica em enviar [troca de tela]
        #peça tem uma segunda tela para preenchimento
        found = browser.find_element_by_css_selector('input[value="Continuar"]').click()
        time.sleep(2)

        #Segunda tela
        select = Select(browser.find_element_by_id('commodityData.chosenProfile.profileID'))
        select.select_by_visible_text('Adicionar nova mercadoria') #add nova mercadoria

        found = browser.find_element_by_id('commodityData.chosenProfile.description')
        found.send_keys(peca) #nome peça

        select = Select(browser.find_element_by_id('commodityData.chosenProfile.unitOfMeasure'))
        select.select_by_visible_text('peças') #unidad de medida

        found = browser.find_element_by_id('commodityData.chosenProfile.quantity')
        found.send_keys(quantidade_peca) #quantidade peça

        found = browser.find_element_by_id('commodityData.chosenProfile.commodityWeight')
        found.send_keys(peso_total) #peso peça

        select = Select(browser.find_element_by_id('commodityData.chosenProfile.manufacturingCountry'))
        time.sleep(1)
        select.select_by_visible_text(pais_coleta) #pais de origem

        found = browser.find_element_by_css_selector('input[value="Adicionar essa mercadoria"]').click() #add a mercadoria
        time.sleep(1)

        browser.find_element_by_xpath(".//*[@id='customsdocuments.CI.chbx']").click() #fatura comercial
        time.sleep(3)

        browser.find_element_by_xpath(".//*[@id='customs.documents.checkbox.letterHeadImage']").click()
        browser.find_element_by_xpath(".//*[@id='customs.documents.checkbox.signatureImage']").click()

        found = browser.find_element_by_css_selector('input[value="Enviar"]').click() #finaliza o pedido
        
        time.sleep(4)

        try:
                time.sleep(3)
                browser.find_element_by_xpath(".//*[@id='customs.documents.checkbox.letterHeadImage']").click()
                time.sleep(1)
                browser.find_element_by_xpath(".//*[@id='customs.documents.checkbox.signatureImage']").click()
                time.sleep(1)
                found = browser.find_element_by_css_selector('input[value="Enviar"]').click() #finaliza o pedido
        
        except:
                pass #falha silenciosamente

        time.sleep(1)
        found = browser.find_element_by_css_selector('input[value="Enviar"]').click() #troca de tela para impressão do AWB
        time.sleep(2)

        found = browser.find_element_by_css_selector('input[value="Imprimir"]').click() # imprime todas as pags do PTA
        time.sleep(5)

        pyautogui.press('enter')  # Preciona tecla enter
                                  # impressora PDF deve estar setada como padrão
        time.sleep(3)
        pyautogui.press('enter')
        pyautogui.hotkey('tab')
        time.sleep(1)
        pyautogui.hotkey('shift','tab')
        
        time.sleep(4)
        pyautogui.typewrite("C:\\Users\\Public\\FEDEX - GM\\"+PTA+'.pdf')
        time.sleep(3)
        pyautogui.press('enter')
        
        time.sleep(10)
        
        found = browser.find_element_by_css_selector('img[alt="labelImage"]') #download da imagem

        time.sleep(2)

        try:
                found.screenshot(PTA + '.png')
        except Exception as ex:
                browser.get_screenshot_as_file(PTA + '.png')

        rgba = Image.open(PTA + '.png')
        rgb = Image.new('RGB', rgba.size, (255, 255, 255))  
        rgb.paste(rgba, mask=rgba.split()[3])               
        rgb.save(PTA + '.pdf', 'PDF', resoultion=100.0)
        os.remove(PTA+'.png')




browser.quit()  # Fecha a janela do navegador
message.Move(tratados) #move o email tratado para a paste de "Tratados"

pyautogui.alert(text='Etiqueta emitida com sucesso', title='Processo GM - FedEx', button='OK') #Janela de OK, confirmando fim do processo

sys.exit() #fecha a aplicação