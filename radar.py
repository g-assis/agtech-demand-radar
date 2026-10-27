import csv,json,re,io,unicodedata,hashlib
from datetime import timedelta
from pathlib import Path
from datetime import datetime,timezone
import requests,pandas as pd
from bs4 import BeautifulSoup
from urllib.parse import urljoin

OUT=Path("output");OUT.mkdir(exist_ok=True)
S=requests.Session();S.headers.update({"User-Agent":"AgTechDemandRadar/0.3 (+public research)"})
LOOKBACK_MONTHS=36
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
def date_column_score(key, kind):
 k=norm(key).replace("_"," ").replace("-"," ")
 if any(x in k for x in ("nascimento","atualizacao","extracao","vencimento","publicacao","inicio vigencia","fim vigencia")):return 0
 event_words=("contrat","assinatura","aprov","formaliza") if kind!="desembolso" else ("libera","desembol","pagamento")
 if any(x in k for x in event_words):return 4 if any(x in k for x in ("data","dt ","date","ano")) else 2
 if re.search(r"\\b(data|date|dt)\\b",k):return 1
 return 0

def parse_date(value):
 if pd.isna(value):return ""
 if isinstance(value,(datetime,pd.Timestamp)):return value.date().isoformat()
 raw=str(value).strip()
 if not raw or raw.lower() in ("nan","nat","none"):return ""
 # ISO, Brazilian date strings, and Excel serial numbers.
 if re.fullmatch(r"\\d{5}(?:\\.0)?",raw):
  try:
   d=pd.to_datetime(float(raw),unit="D",origin="1899-12-30")
   if 1990<=d.year<=datetime.now(timezone.utc).year+1:return d.date().isoformat()
  except (ValueError,OverflowError,TypeError):pass
 try:
  parsed=pd.to_datetime(raw,dayfirst=True,errors="coerce")
  if pd.notna(parsed) and 1990<=parsed.year<=datetime.now(timezone.utc).year+1:return parsed.date().isoformat()
 except (ValueError,OverflowError,TypeError):pass
 return ""

def extract(df,source,url,kind="financiamento"):
 result=[];picked={}
 columns=[str(c) for c in df.columns]
 ranked=sorted(columns,key=lambda c:date_column_score(c,kind),reverse=True)
 eligible=[c for c in ranked if date_column_score(c,kind)>0]
 for col in eligible[:8]:
  series=df[col].dropna().head(80)
  valid=sum(bool(parse_date(v)) for v in series)
  if valid:picked[col]=valid
 print(json.dumps({"fonte":source,"arquivo":url.split("#")[0].split("/")[-1],"aba":url.split("#sheet=")[-1] if "#sheet=" in url else "CSV","colunas":columns[:65],"datas_candidatas":picked},ensure_ascii=False))
 for _,x in df.iterrows():
  d={norm(k):str(v) for k,v in x.items() if pd.notna(v)}
  desc=" | ".join(f"{k}: {v}" for k,v in d.items() if any(t in k for t in ("projeto","setor","atividade","produto","finalidade","objetivo","tema","descricao")))
  org=next((v for k,v in d.items() if any(t in k for t in ("cliente","empresa","beneficiario","proponente"))),"")
  val=next((v for k,v in d.items() if "valor" in k),"")
  if not desc:desc=" | ".join(f"{k}: {v}" for k,v in d.items() if len(v)>3)[:1200]
  a=row(source,url,desc[:240] or "Operacao",desc,org,val,kind=kind)
  if not a:continue
  date=""
  for col in eligible:
   if col in x.index:
    date=parse_date(x[col])
    if date:break
  a["data_origem"]=date
  a["coluna_data"]=col if date else ""
  result.append(a)
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
 ]
 results=[];errors=[]
 for filename,kind in datasets:
  url=base+filename
  try:
   raw=get(url).content
   if not raw.startswith(b"PK"):raise ValueError("Resposta nao e XLSX valido")
   workbook=pd.ExcelFile(io.BytesIO(raw))
   inspected=0;matched=0
   for sheet in workbook.sheet_names:
    if norm(sheet).startswith(("guia","legenda","crd_-_condicoes")):continue
    df=pd.read_excel(workbook,sheet_name=sheet,dtype=str)
    inspected+=len(df)
    parsed=extract(df,"Finep",url+"#sheet="+sheet,kind=kind)
    for item in parsed:item["tipo_evidencia"]=kind
    matched+=len(parsed);results.extend(parsed)
   print(json.dumps({"fonte":"Finep","arquivo":filename,"abas":workbook.sheet_names,"linhas_lidas":inspected,"linhas_agro":matched},ensure_ascii=False))
   if inspected==0:errors.append(f"{filename}: planilha vazia")
   elif matched==0:errors.append(f"{filename}: nenhuma linha classificada como agro em {inspected} linhas")
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
def clean_org(s):
 x=norm(s)
 x=re.sub(r"\\b(s/?a|ltda|eireli|sa|me|epp)\\b"," ",x)
 return re.sub(r"[^a-z0-9]","",x)[:100]

def find_date(text):
 # Conservative: only dates explicitly labeled in published records.
 patterns=[r"(?:data(?:\s+de)?\s+(?:contratacao|aprovacao|liberacao|desembolso|assinatura))[:\s-]{0,8}(\d{1,2}/\d{1,2}/\d{4})",
           r"(?:data(?:\s+de)?\s+(?:contratacao|aprovacao|liberacao|desembolso|assinatura))[:\s-]{0,8}(\d{4}-\d{2}-\d{2})"]
 from datetime import datetime as dt
 for pattern in patterns:
  m=re.search(pattern,norm(text))
  if m:
   try:return dt.strptime(m.group(1),"%d/%m/%Y" if "/" in m.group(1) else "%Y-%m-%d").date().isoformat()
   except ValueError:pass
 return ""

def problem(text):
 t=norm(text)
 matched=[k for k,words in TOPICS.items() if any(w in t for w in words)]
 return matched or ["nao_identificado"]

def normalize_records(raw):
 # Exclude rejected financing; disbursement is a separate financial event.
 cutoff=(datetime.now(timezone.utc)-timedelta(days=LOOKBACK_MONTHS*30.44)).date().isoformat()
 collected={}
 excluded_old=0; missing_dates=0
 for item in raw:
  if item["tipo_evidencia"]=="operacao_nao_aprovada":continue
  date=item.get("data_origem") or find_date(item["descricao"])
  if not date: missing_dates+=1
  elif date<cutoff:
   excluded_old+=1;continue
  org=clean_org(item["organizacao"])
  # Same evidence type and textual record in separate exports is one signal.
  key=hashlib.sha256(("|".join([item["fonte"],item["tipo_evidencia"],org,norm(item["descricao"])[:400]])).encode()).hexdigest()[:16]
  item.update(data_evidencia=date,periodo="ultimos_36_meses" if date else "data_nao_verificada",id_evidencia=key,problema_economico="; ".join(problem(item["descricao"])),evidencia_comercial="investimento_financiado" if item["tipo_evidencia"] in ("operacao_contratada","financiamento") else "desembolso_reportado",confianca_data="coluna_estruturada" if item.get("data_origem") else ("extraida_do_texto" if date else "nao_verificada"))
  collected[key]=item
 return list(collected.values()),dict(excluidos_antigos=excluded_old,sem_data_verificavel=missing_dates,janela_meses=LOOKBACK_MONTHS)

def main():
 raw=[];status=[]
 for name,fn in [("BNDES",bndes),("Finep",finep)]:
  try:
   data=fn();raw+=data;status.append(dict(fonte=name,estado="ok" if data else "sem_registros",registros=len(data)))
  except Exception as e:status.append(dict(fonte=name,estado="erro",registros=0,motivo=str(e)[:500]))
 allrows,quality=normalize_records(raw)
 cols=["id_evidencia","fonte","tipo_evidencia","organizacao","titulo","descricao","valor_reportado","temas","data_evidencia","periodo","problema_economico","evidencia_comercial","confianca_data","url_fonte","demanda_comercial_comprovada","data_origem","coluna_data"]
 with (OUT/"evidencias.csv").open("w",newline="",encoding="utf-8-sig") as f:
  w=csv.DictWriter(f,cols);w.writeheader();w.writerows(allrows)
 groups={}
 for a in allrows:
  if a["periodo"]!="ultimos_36_meses":continue
  for tag in a["problema_economico"].split("; "):
   g=groups.setdefault(tag,dict(tema=tag,n_registros=0,fontes=set(),organizacoes=set(),contratacoes=0,desembolsos=0))
   g["n_registros"]+=1;g["fontes"].add(a["fonte"])
   if a["organizacao"]:g["organizacoes"].add(clean_org(a["organizacao"]))
   if a["tipo_evidencia"]=="desembolso":g["desembolsos"]+=1
   else:g["contratacoes"]+=1
 summary=[dict(tema=v["tema"],n_registros=v["n_registros"],fontes=", ".join(sorted(v["fontes"])),n_organizacoes=len(v["organizacoes"]),contratacoes=v["contratacoes"],desembolsos=v["desembolsos"]) for v in groups.values()]
 summary.sort(key=lambda x:-x["n_registros"])
 with (OUT/"temas.csv").open("w",newline="",encoding="utf-8") as f:
  w=csv.DictWriter(f,["tema","n_registros","fontes","n_organizacoes","contratacoes","desembolsos"]);w.writeheader();w.writerows(summary)
 (OUT/"evidencias.json").write_text(json.dumps(allrows,ensure_ascii=False),encoding="utf-8")
 (OUT/"temas.json").write_text(json.dumps(summary,ensure_ascii=False),encoding="utf-8")
 state=dict(gerado_em_utc=datetime.now(timezone.utc).isoformat(),total_registros=len(allrows),registros_recentes=sum(x["periodo"]=="ultimos_36_meses" for x in allrows),registros_brutos=len(raw),fontes=status,qualidade=quality,nota="Registros sem data ficam separados e nao entram nos temas recentes. Investimento nao comprova demanda comercial.")
 (OUT/"status.json").write_text(json.dumps(state,ensure_ascii=False,indent=2),encoding="utf-8")
 print(json.dumps(state,ensure_ascii=False,indent=2))
 if not allrows:raise SystemExit("Sem evidencias coletadas; revisar fontes")
if __name__=="__main__":main()
