import sys,unittest,tempfile
from pathlib import Path
from datetime import date
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'src'))
from analisador import cents,dt,stats,status,business_add,business_roll,analyze
from openpyxl import Workbook

class Rules(unittest.TestCase):
    def test_centavos(self):
        self.assertEqual(cents(657.9399999999996),65794)
        self.assertEqual(cents('1.234,56'),123456)
        self.assertIsNone(cents(None));self.assertEqual(cents(0),0)
    def test_feriado_e_fim_semana(self):
        self.assertEqual(business_add(date(2026,1,30),3,{date(2026,2,2)}),date(2026,2,5))
        self.assertEqual(business_roll(date(2026,2,1),{date(2026,2,2)}),date(2026,2,3))
    def test_ausencia_nao_vira_zero(self):
        self.assertIsNone(dt('gasto total'));self.assertIsNone(stats([])['media'])
        self.assertEqual(stats([0,10])['media'],5)
        self.assertEqual(status(None),'Não informado');self.assertEqual(status(' Ressalva '),'Ressalva')
    def test_populacao_e_datas_contraditorias(self):
        with tempfile.TemporaryDirectory() as tmp:
            root=Path(tmp);wb=Workbook();ws=wb.active;ws.title='Dados'
            headers={'A':'nº processo SEI','B':'SERVIDOR','C':'SECRETARIA','D':'nº nota','E':'valor adt','F':'valor utl','G':'valor dev','I':'DATA PAGTO.','W':'Tipo','Z':'EXAME SMF','AA':'EXAME CGM'}
            for c,v in headers.items():ws[c+'2']=v
            for i,note,kind,use,ret in [(3,1,'serviço',100,0),(4,2,'consumo',0,100)]:
                for c,v in {'A':'SEI1','B':'Servidor A','C':'SMS','D':note,'E':100,'F':use,'G':ret,'H':'Jan','I':date(2026,1,10),'K':date(2026,2,9),'L':date(2026,2,12),'M':date(2026,2,19),'O':date(2026,1,11) if use else 'sem uso','P':date(2026,1,9) if ret else 'gasto total','Q':date(2026,1,15),'W':kind,'Z':'aprovado','AA':'aprovado'}.items():ws[f'{c}{i}']=v
            ws['S50']='=IF(A50="","","Não")'
            wb.save(root/'fonte.xlsx')
            cfg={'arquivo':'fonte.xlsx','aba':'Dados','inicio':'2026-01-01','fim':'2026-06-30','data_observacao':'2026-11-06','feriados':[]}
            rows,tables,summary,meta=analyze(cfg,root)
            self.assertEqual(summary['registros'],2);self.assertEqual(summary['processos_unicos'],1)
            self.assertEqual(summary['adiantado'],200);self.assertEqual(summary['diferenca_financeira'],0)
            self.assertIsNone(rows[1]['dias_pagamento_devolucao'])
            self.assertEqual(tables['tempos'][1]['n'],0)
            self.assertTrue(any(r['codigo']=='EVENTO_ANTES_PAGAMENTO' for r in tables['alertas']))

if __name__=='__main__':unittest.main()
