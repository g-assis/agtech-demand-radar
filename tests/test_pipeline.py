import unittest
from datetime import datetime,timezone
import radar

class RadarTests(unittest.TestCase):
 def test_normalizacao_organizacao(self):
  self.assertEqual(radar.clean_org("Cooperativa Agrícola Ltda."),radar.clean_org("COOPERATIVA AGRICOLA LTDA"))
 def test_data_identificavel(self):
  self.assertEqual(radar.find_date("data de contratacao: 19/10/2026"),"2026-10-19")
 def test_nao_inferir_data(self):
  self.assertEqual(radar.find_date("Custo estimado 2026"),"")
 def test_sem_demanda_comprovada(self):
  a=radar.row("Finep","https://example.org","Projeto de armazenagem de graos","silos e secagem")
  self.assertEqual(a["demanda_comercial_comprovada"],"NAO")
 def test_rejeitado_excluido(self):
  a=radar.row("Finep","https://example.org","Projeto de armazenagem de graos","data de contratacao: 01/10/2026 silo")
  a["tipo_evidencia"]="operacao_nao_aprovada"
  out,_=radar.normalize_records([a]);self.assertEqual(out,[])
 def test_contratacao_e_desembolso_nao_somados_como_projetos(self):
  a=radar.row("Finep","https://example.org","Silo e graos","data de contratacao: 01/10/2026 silos")
  a["tipo_evidencia"]="operacao_contratada"
  b=dict(a);b["tipo_evidencia"]="desembolso"
  out,_=radar.normalize_records([a,b])
  self.assertEqual(len(out),2)
if __name__=="__main__": unittest.main()
