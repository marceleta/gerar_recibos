from kivymd.app import MDApp
from kivymd.uix.widget import MDWidget
from kivymd.uix.list import OneLineListItem
from kivy.core.window import Window
from kivymd.uix.screen import MDScreen
from kivymd.uix.screenmanager import MDScreenManager
from kivy.config import Config
from kivy.properties import ListProperty
import os, sys
import datetime
from docx import Document
from docx.shared import Pt


Config.set('graphics', 'resizable', False) 


class SelecaoClientes(MDScreen):
    
    def __init__(self, **kwargs):
        super(SelecaoClientes, self).__init__(**kwargs)
        self.lista_clientes = self.ids.lista_clientes
        self.lista_selecionados = self.ids.lista_selecionados
        
        self.construir_lista()
    
    def construir_lista(self):
        
        clientes = get_clientes()
        for cliente in clientes:
            self.lista_clientes.add_widget(OneLineListItem(text=cliente, 
                                                           on_release=self.item_escolhido_cliente))
            
    
        
    def item_escolhido_cliente(self, item_lista):
        clientes = self.clientes_para_selecionar().copy()
        
        if item_lista in clientes:
            self.remover_item_clientes(item_lista)
            self.adicionar_item_selecionados(item_lista)
            
    
    def item_escolhido_selecionados(self, item_lista):
        clientes = self.lista_selecionados.children
        
        if item_lista in clientes:
            self.remover_item_selecionados(item_lista)
            self.adicionar_item_clientes(item_lista)
                                    
    
    def adicionar_todos_clientes(self):
        clientes = self.lista_clientes.children.copy()
        
        for cliente in clientes:
            self.adicionar_item_selecionados(cliente)
            self.remover_item_clientes(cliente)
    
    def remover_todos_selecionados(self):
        clientes = self.clientes_selecionados().copy()
        
        for cliente in clientes:
            self.adicionar_item_clientes(cliente)
            self.remover_item_selecionados(cliente)
        
    
    def adicionar_item_clientes(self, item):
        novo_item = OneLineListItem(text=item.text, on_release=self.item_escolhido_cliente)
        self.lista_clientes.add_widget(novo_item)        
    
    def remover_item_clientes(self, item):
        self.lista_clientes.remove_widget(item)
        
        
    def adicionar_item_selecionados(self, item_lista):
        novo_item = OneLineListItem(text=item_lista.text, on_release=self.item_escolhido_selecionados)
        self.lista_selecionados.add_widget(novo_item)
        
    def remover_item_selecionados(self, item_lista):
        self.lista_selecionados.remove_widget(item_lista)
        
        
    def clientes_selecionados(self):
        return self.lista_selecionados.children
    
    def clientes_para_selecionar(self):
        return self.lista_clientes.children
    
    def get_clientes(self):
        pastas = os.walk('clientes')
        pastas = next(pastas)[1]
        
        return pastas
    
    def cancelar(self):
        sys.exit(0)
        
    def tela_gerar_boletos(self):
        
        qtd_clientes_selecionados = len(self.clientes_selecionados())
        
        if qtd_clientes_selecionados > 0:
            self.manager.get_screen('gerar_recibo').clientes = self.clientes_selecionados()
            self.manager.get_screen('gerar_recibo').proximo_cliente()
            self.parent.current = 'gerar_recibo'
        
        
class GerarRecibo(MDScreen):
    
    def __init__(self, **kwargs):
        super(GerarRecibo, self).__init__(**kwargs)
        self.clientes = []
        self.num_clientes = 0
        self.cliente_em_producao = 0
        
    def proximo_cliente(self):
        self.limpar_labels()
        self.num_clientes = len(self.clientes)
        if self.num_clientes > 0:            
            self.ids.lb_num_cliente.text = "{}/{}".format(str(self.cliente_em_producao + 1), str(len(self.clientes)))
            self.ids.lb_nome_cliente.text = str(self.clientes[self.cliente_em_producao].text)
            self.cliente_em_producao = self.cliente_em_producao + 1
        
            
    def limpar_labels(self):
        self.ids.tf_valor_parcela.text = ''
        self.ids.tf_data_inicio_contrato.text = ''
        self.ids.tf_data_final_contrato.text = ''
        self.ids.tf_valor_extenso.text = ''
    
    def pegar_texto_labels (self):
        
        valores = {}
        valores['valor_parcela'] = self.ids.tf_valor_parcela.text
        valores['data_inicio_contrato'] = self.ids.tf_data_inicio_contrato.text
        valores['data_final_contrato'] = self.ids.tf_data_final_contrato.text
        valores['valor_extenso'] = self.ids.tf_valor_extenso.text
        
        return valores
    
    
    def gerar_documento(self):
        valores = self.pegar_texto_labels()
        
        
    def validar_valor_parcela(self, instance, input):
        print(input)
        ''' if input == '\t':
            print(input) '''
            
    def tratar_tabulacao(self, instance, texto):
        tabulado = False
        if len(texto) > 0:
            if texto[-1] == '\t':
                texto = texto[:-1]
                instance.text = texto
                tabulado = True
        return tabulado, texto
            
    def on_text_valor(self, instance, input):

        if len(input) > 0:
            if input[-1] == '\t':
                input = input[:-1]
                instance.text = input
                self.ids.tf_data_inicio_contrato.focus = True
                
    def on_text_data_inicio(self, instance, input):
        tabulado, texto = self.tratar_tabulacao(instance, input)
        
        if tabulado:
            self.ids.tf_data_final_contrato.focus = True
            
    def on_text_data_final(self, instance, input):
        tabulado, texto = self.tratar_tabulacao(instance, input)
        
        if tabulado:
            self.ids.tf_valor_extenso.focus = True
            
    def on_text_valor_extenso(self, instance, input):
        print(input)
        tabulado, texto = self.tratar_tabulacao(instance, input)
        
        if tabulado:
            self.ids.btn_gerar_recibo.focus = True
                
        
    def cancelar(self):
        self.clientes = []
        self.cliente_em_producao = 0
        self.parent.current = 'selecao_cliente'


class RecibosApp(MDApp):
    title = "Gerador de Recibos"
    Window.size = (700, 500)
    
    
    def build(self):
        self.theme_cls.primary_palette = "Blue"
        sm = MDScreenManager()
        sm.add_widget(SelecaoClientes(name='selecao_cliente'))
        sm.add_widget(GerarRecibo(name='gerar_recibo'))       
        
        return sm
    
    def on_start(self):
        pass


class CalculoRecibo():
    
    @classmethod
    def dados_entrada(cls, valores_iniciais):
        valor = valores_iniciais['valor']
        valor_extenso = valores_iniciais['v_extenso']
        dt_inicio_contrato = valores_iniciais['data_inicio_contrato']
        dt_final_contrato = valores_iniciais['data_final_contrato']
        
        
        lista_recibos = []
        mes = 0
        for periodo in range(12):
            
            
            dados_recibo = {}
            
            if periodo == 0:
                dados_recibo['mes'] = periodo + 1
                dados_recibo['dt_vencimento'] = cls.soma_mes(dt_inicio_contrato, 0)
                dados_recibo['mes_referencia'] = cls.mes_referencia(dados_recibo['dt_vencimento'])
                dados_recibo['periodo_ref'] = cls.periodo_ref(dados_recibo['dt_vencimento'])
                
            else:
                mes = mes + 1
                dados_recibo['mes'] = mes + 1
                
                dados_recibo['dt_vencimento'] = cls.soma_mes(dt_inicio_contrato, mes)
                dados_recibo['mes_referencia'] = cls.mes_referencia(dados_recibo['dt_vencimento'], mes)
                dados_recibo['periodo_ref'] = cls.periodo_ref(dados_recibo['dt_vencimento'])
                
            dados_recibo['info_contrato'] = 'Inicio {} Vencimento {}'.format(dt_inicio_contrato, dt_final_contrato)
                
                
            
            lista_recibos.append(dados_recibo)
                
        return lista_recibos
                
                         
    @classmethod
    def soma_mes(cls, data, numero_mes):
        
        split = data.split('/')
        mes = int(split[1])
        mes = mes + numero_mes
        
        if mes > 12:
            ano = int(split[2])
            ano = ano + 1
            mes = mes - 12
            nova_data = split[0] + '/' + str(mes) + '/' + str(ano)
        else:
            nova_data = split[0] + '/' + str(mes) + '/' + split[2]
            
        return nova_data
    
    @classmethod
    def mes_referencia(cls, data, novo_mes=None):
        
        split = data.split('/')
        
        mes = int(split[1])
        ano = int(split[2])
        
        if novo_mes != None:
            if mes > 12:
                mes = mes - 12
                ano = ano + 1
            m_referencia = str(mes) + '/' + str(ano)
        else:
            m_referencia = split[1] + '/' + split[2]
        
        return m_referencia
    
    @classmethod
    def periodo_ref(cls, data_vencimento):
        
        split = data_vencimento.split('/')
        dia = int(split[0]) - 1
        mes = int(split[1]) + 1
        ano = int(split[2])
        
        
        if mes > 12:
            mes = mes - 12
            ano = ano + 1
        
        data_ref_final = "{}/{}/{}".format(str(dia), str(mes), str(ano))
        
        return "{} à {}".format(data_vencimento, data_ref_final)
                
    #Soma n aos meses sem alterar a             
    @classmethod
    def soma_mes(cls, data, numero_mes):
        
        split = data.split('/')
        mes = int(split[1])
        mes = mes + numero_mes
        
        if mes > 12:
            ano = int(split[2])
            ano = ano + 1
            mes = numero_mes
            nova_data = split[0] + '/' + mes + '/' + ano
        else:
            nova_data = split[0] + '/' + mes + '/' + split[2]
            
        return nova_data
    
    @classmethod
    def mes_referencia(cls, data):
        
        split = data.split('/')
        
        m_referencia = split[1] + '/' + split[2]
        
        return m_referencia
    
    @classmethod
    def periodo_ref(cls, data_vencimento):
        
        split = data_vencimento.split('/')
        
        
class Recibos():
    
    @staticmethod
    def gerar(modelo, calculo_recibo):
        pass
    
    


def get_clientes():
    clientes = os.walk('clientes')
    clientes = next(clientes)[1]
    
    return clientes
        


if __name__ == '__main__':
    RecibosApp().run()
