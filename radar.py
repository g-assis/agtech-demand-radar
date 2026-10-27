import csv,json,re,io,unicodedata
from pathlib import Path
from datetime import datetime,timezone
import requests,pandas as pd
from bs4 import BeautifulSoup
from urllib.parse import urljoin

OUT=Path("output");OUT.mkdir(exist_ok=True)
S=requests.Session();S.headers.update({"User-Agent":"AgTechDemandRadar/0.3 (+public research)"})
KEYS=["agro","agric","cooperativ","grao","silo","armazen","sement","alimento","frigor","latic","irrig","pecuar","safra","fertiliz","colheita","cafe","sucro","processamento","residuo"]
TOPICS={"armazenagem":["silo","armazen","secagem"],"processamento":["processamento","beneficiamento","agroindustr"],"qualidade":["qualidade","inspec","rastreab"],"energia_agua":["energia","irrig","agua"],"residuos":["residuo","reaproveitamento","subproduto"],"automacao":["automat","sensor","iot","digital","inteligencia artificial"]}
def norm(v):
 return re.sub(r"\s+"," ",unicodedata.normalize("NFKD",str(v or "")).encode("ascii","ignore").decode().lower())
def get(url):
 r=S.get(url,timeout=55);r.raise_for_status();return r
def readtab(url):
 data=get(url).content
 if len(data)>90000000:raise ValueError("Arquivo >90MB")
 if data[:2]==b"PK" or url.lower().split("?")[0].endswith((".xls",".xlsx")):
  return pd.read_excel(io.BytesIO(data),dtype=str)
 for enc in ("utf-8-sig","cp1252","latin1"):
  try: txt=data.decode(enc);break
  except UnicodeDecodeError:continue
 import csv as _csv
 try:sep=_csv.Sniffer().sniff(txt[:8000],delimiters=";,\t").delimiter
 except _csv.Error:sep=","
 return pd.read_csv(io.StringIO(txt),sep=sep,dtype=str,on_bad_lines="skip")
def row(source,url,title,detail,org="",value="",kind="financiamento"):
 full=norm(" ".join([title,detail,org]))
 if not any(k in full for k in KEYS):return None
 tags=[k for k,v in TOPICS.items() if any(w in full for w in v)]
 return dict(fonte=source,tipo_evidencia=kind,organizacao=org,titulo=title[:240],descricao=detail[:1200],valor_reportado=str(value),temas="; ".join(tags) or "nao_classificado",url_fonte=url,demanda_comercial_comprovada="NAO")
def extract(df,source,url):
 result=[]
 for _,x in df.iterrows():
  d={norm(k):str(v) for k,v in x.items() if pd.notna(v)}
  desc=" | ".join(f"{k}: {v}" for k,v in d.items() if any(t in k for t in ("projeto","setor","atividade","produto","finalidade","objetivo","tema","descricao")))
  org=next((v for k,v in d.items() if any(t in k for t in ("cliente","empresa","beneficiario","proponente"))),"")
  val=next((v for k,v in d.items() if "valor" in k),"")
  if not desc: desc=" | ".join(f"{k}: {v}" for k,v in d.items() if len(v)>3)[:1200]
  a=row(source,url,desc[:240] or "Operacao",desc,org,val)
  if a:result.append(a)
 return result
def bndes():
 url="https://dadosabertos.bndes.gov.br/api/3/action/package_show?id=operacoes-financiamento"
 resources=get(url).json()["result"]["resources"]
 candidates=[r["url"] for r in resources if "csv" in str(r.get("format","")).lower() and r.get("url")]
 if not candidates:raise ValueError("Nenhum recurso CSV BNDES")
 results=[];errs=[]
 for link in candidates[:3]:
  try:results+=extract(readtab(link),"BNDES",link)
  except Exception as exc:errs.append(str(exc))
 if not results and errs:raise RuntimeError("; ".join(errs)[:350])
 return results
def finep():
 # Official XLSX downloads linked from the Finep legacy downloads portal.
 # Direct URLs are stable file names; weekly refreshes overwrite file contents.
 base="https://download.finep.gov.br/"
 datasets=[
  ("Contratacao.xlsx","operacao_contratada"),
  ("Liberacao.xlsx","desembolso"),
  ("Nao_Aprovados.xlsx","operacao_nao_aprovada"),
 ]
 results=[];errors=[]
 for filename,kind in datasets:
  url=base+filename
  try:
   df=readtab(url)
   # Some public exports use nonstandard column names; retain the row text
   # for searching when none of the expected descriptive headers exist.
   parsed=extract(df,"Finep",url)
   for item in parsed:item["tipo_evidencia"]=kind
   results.extend(parsed)
  except Exception as exc:errors.append(f"{filename}: {str(exc)[:180]}")
 if errors:raise RuntimeError("Falha parcial na Finep: "+"; ".join(errors))
 return results
def embrapii():
 url="https://embrapii.org.br/projetos/"
 soup=BeautifulSoup(get(url).text,"html.parser");result=[];seen=set()
 for a in soup.select("a[href]"):
  href=urljoin(url,a["href"]);title=a.get_text(" ",strip=True)
  if href in seen or len(title)<12 or "/projetos/" not in href or href.rstrip("/")==url.rstrip("/"):continue
  seen.add(href);hit=row("Embrapii",href,title,title,kind="projeto_pesquisa")
  if hit:result.append(hit)
 return result
def main():
 allrows=[];status=[]
 for name,fn in [("BNDES",bndes),("Finep",finep),("Embrapii",embrapii)]:
  try:
   data=fn();allrows+=data;status.append(dict(fonte=name,estado="ok" if data else "sem_registros",registros=len(data)))
  except Exception as e:status.append(dict(fonte=name,estado="erro",registros=0,motivo=str(e)[:500]))
 allrows=list({(a["fonte"],a["url_fonte"],a["titulo"],a["organizacao"]):a for a in allrows}.values())
 cols=["fonte","tipo_evidencia","organizacao","titulo","descricao","valor_reportado","temas","url_fonte","demanda_comercial_comprovada"]
 with (OUT/"evidencias.csv").open("w",newline="",encoding="utf-8-sig") as f:
  w=csv.DictWriter(f,cols);w.writeheader();w.writerows(allrows)
 groups={}
 for a in allrows:
  for tag in a["temas"].split("; "):
   g=groups.setdefault(tag,dict(tema=tag,n_registros=0,fontes=set(),organizacoes=set()))
   g["n_registros"]+=1;g["fontes"].add(a["fonte"])
   if a["organizacao"]:g["organizacoes"].add(a["organizacao"])
 summary=[dict(tema=v["tema"],n_registros=v["n_registros"],fontes=", ".join(sorted(v["fontes"])),n_organizacoes=len(v["organizacoes"])) for v in groups.values()]
 summary.sort(key=lambda x:-x["n_registros"])
 with (OUT/"temas.csv").open("w",newline="",encoding="utf-8") as f:
  w=csv.DictWriter(f,["tema","n_registros","fontes","n_organizacoes"]);w.writeheader();w.writerows(summary)
 (OUT/"evidencias.json").write_text(json.dumps(allrows,ensure_ascii=False),encoding="utf-8")
 (OUT/"temas.json").write_text(json.dumps(summary,ensure_ascii=False),encoding="utf-8")
 state=dict(gerado_em_utc=datetime.now(timezone.utc).isoformat(),total_registros=len(allrows),fontes=status,nota="Financiamentos e projetos sao sinais de investimento, NAO demanda comercial comprovada.")
 (OUT/"status.json").write_text(json.dumps(state,ensure_ascii=False,indent=2),encoding="utf-8")
 print(json.dumps(state,ensure_ascii=False,indent=2))
 if not allrows:raise SystemExit("Nenhuma evidencia coletada; verifique status.json")
if __name__=="__main__":main()
