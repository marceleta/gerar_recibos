import unittest
from doc_recibo import CalculoRecibo, Recibos
from docx import Document

class CalculoReciboTest(unittest.TestCase):
    
    def setUp(self):
        dados_iniciais = {
                'valor': '100.00',
                'v_extenso': 'Cem reais',
                'data_inicio_contrato': '20/01/2023',
                'data_final_contrato':'19/01/2024'}
        
        self.calculo_recibo = CalculoRecibo.gerar(dados_iniciais)
        
        
    def test_tipo_de_retorno(self):
        
        esperado = []
        
        self.assertEqual(type(self.calculo_recibo), type(esperado))
        
    def test_quantidade_de_meses(self):
        resultado = len(self.calculo_recibo)
        esperado = 12
        
        self.assertEqual(resultado, esperado)
        
    def test_dados_gerados_vencimento(self):
        resultado = self.calculo_recibo[0]['{VENCIMENTO}']
        esperado = '20/1/2023'
        
        self.assertEqual(resultado, esperado)
        
    def test_dados_gerados_mes_referencia(self):
        resultado = self.calculo_recibo[0]['{MES_REFER}']
        esperado = '1/2023'
        
        self.assertEqual(resultado, esperado)
        
    def test_dados_gerados_periodo(self):
        resultado = self.calculo_recibo[0]['{PERIODO}']
        esperado = '20/1/2023 à 19/2/2023'
        
        self.assertEqual(resultado, esperado)
        
class RecibosTest(unittest.TestCase):
    
    def setUp(self):
        path_cliente = 'clientes/ADONIAS'
        dados_iniciais = {
                'valor': '100.00',
                'v_extenso': 'Cem reais',
                'data_inicio_contrato': '20/01/2023',
                'data_final_contrato':'19/01/2024'}
        
        calculo_recibo = CalculoRecibo.gerar(dados_iniciais)
        
        self.recibos = Recibos(path_cliente, calculo_recibo)
        
    def test_tipo_de_documento(self):
        
        resultado = type(self.recibos.criar_doc())
        esperado = type(Document())
        
        self.assertEqual(esperado, resultado)


if __name__ =='__main__':
    unittest.main()
        