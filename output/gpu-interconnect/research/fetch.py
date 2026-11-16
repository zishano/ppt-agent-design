from pathlib import Path
import requests,concurrent.futures,re,json
from bs4 import BeautifulSoup
R=Path(__file__).resolve().parent
urls={'S1':'https://www.nvidia.com/en-us/data-center/nvlink/','S2':'https://www.nvidia.com/en-us/data-center/gb200-nvl72/','S3':'https://ualinkconsortium.org/','S4':'https://www.amd.com/en/products/accelerators/instinct/mi300/mi300x.html','S5':'https://rocm.docs.amd.com/en/latest/conceptual/gpu-arch/mi300.html','S6':'https://www.amd.com/en/technologies/infinity-architecture.html'}
def get(item):
 k,u=item
 try:
  r=requests.get(u,timeout=35);r.raise_for_status();r.encoding='utf-8';s=BeautifulSoup(r.text,'html.parser')
  for x in s(['script','style','nav','footer']):x.decompose()
  t=s.get_text(' ',strip=True);(R/(k+'.txt')).write_text(t);(R/(k+'.url')).write_text(u)
  out=[]
  for term in ['200G','1024','1,024','1.0','2.0','1.8','900','896','Infinity','NVSwitch','load','store']:
   m=re.search(re.escape(term),t,re.I)
   if m:out.append(t[max(0,m.start()-110):m.end()+270])
  return k,out
 except Exception as e:return k,str(e)
if __name__=='__main__':
 with concurrent.futures.ThreadPoolExecutor(max_workers=6) as ex:
  for o in ex.map(get,urls.items()):print(o)
