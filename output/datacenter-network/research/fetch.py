from pathlib import Path
import requests,concurrent.futures,re,json
from bs4 import BeautifulSoup
R=Path(__file__).resolve().parent
urls={'S1':'https://ethernetalliance.org/technology/ethernet-roadmap/','S2':'https://www.infinibandta.org/about-infiniband/','S3':'https://www.infinibandta.org/roce/','S4':'https://ultraethernet.org/','S5':'https://ultraethernet.org/uec-specification/','S6':'https://docs.nvidia.com/networking/display/mlnxenv24102180lts/rdma+over+converged+ethernet+(roce)'}
def get(item):
 k,u=item
 try:
  r=requests.get(u,timeout=35);r.raise_for_status();r.encoding='utf-8';s=BeautifulSoup(r.text,'html.parser')
  for x in s(['script','style','nav','footer']):x.decompose()
  t=s.get_text(' ',strip=True);t='\n'.join(x.rstrip() for x in t.splitlines())+'\n';(R/(k+'.txt')).write_text(t);(R/(k+'.url')).write_text(r.url)
  terms=['1.0','1.1','2026','UDP','routable','PFC','credit','RDMA','transport','lossless']
  matches=[]
  for term in terms:
   m=re.search(re.escape(term),t,re.I)
   if m:matches.append(t[max(0,m.start()-110):m.end()+230])
  return k,r.status_code,matches
 except Exception as e:return k,str(e)
if __name__=='__main__':
 with concurrent.futures.ThreadPoolExecutor(max_workers=6) as ex:
  for result in ex.map(get,urls.items()):print(result)
