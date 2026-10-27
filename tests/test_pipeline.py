import unittest
from datetime import datetime,timezone
import radar

class RadarTests(unittest.TestCase):
 def test_normalizacao_organizacao(self):
  self.assertEqual(radar.clean_org("Cooperativa Agrícola Ltda."),radar.clean_org("COOPERATIVA AGRICOLA LTDA"))
 def test_data_identificavel(self):
  self.assertEqual(radar.find_date("data de contratacao: 19/10/2026"),"2026-10-19")
 def test_data_excel_serial(self):
  self.assertEqual(radar.parse_date("45292"),"2024-01-01")
 def test_date_column_mapping(self):
  import pandas as pd
  df=pd.DataFrame({"Data da Contratação":["19/10/2026"],"Descrição do Projeto":["modernizacao de silos"],"Empresa":["Cooperativa ABC"]})
  rows=radar.extract(df,"Finep","https://example.org/Contratacao.xlsx",kind="operacao_contratada")
  self.assertEqual(len(rows),1)
  self.assertEqual(rows[0]["data_origem"],"2026-10-19")
  normalized,_=radar.normalize_records(rows)
  self.assertEqual(normalized[0]["periodo"],"ultimos_36_meses")
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
