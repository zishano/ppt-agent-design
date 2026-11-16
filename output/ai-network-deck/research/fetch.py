import requests,concurrent.futures
from bs4 import BeautifulSoup
from pathlib import Path
R=Path(__file__).parent
urls={
'S1':'https://docs.nvidia.com/deeplearning/nccl/user-guide/docs/usage/collectives.html',
'S2':'https://www.nvidia.com/en-us/data-center/gb200-nvl72/',
'S3':'https://www.nvidia.com/en-us/networking/products/infiniband/quantum-2/',
'S4':'https://www.nvidia.com/en-us/networking/spectrum-x/',
'S5':'https://docs.nvidia.com/megatron-core/developer-guide/latest/api-guide/tensor_parallel.html',
'S6':'https://docs.nvidia.com/deeplearning/nccl/user-guide/docs/usage/communicators.html'}
def get(it):
 k,u=it
 try:
  r=requests.get(u,timeout=40);r.raise_for_status();s=BeautifulSoup(r.text,'html.parser')
  for x in s(['script','style','nav','footer']):x.decompose()
  t=s.get_text(' ',strip=True);(R/(k+'.txt')).write_text(t);(R/(k+'.url')).write_text(u)
  terms=['allreduce','all-reduce','alltoall','72','130','1.8','400','rdma','adaptive','tensor']
  import re
  matches=[t[max(0,m.start()-130):m.end()+240] for m in re.finditer('|'.join(re.escape(a) for a in terms),t,re.I)]
  return k,r.status_code,matches[:18]
 except Exception as e:return k,str(e)
with concurrent.futures.ThreadPoolExecutor(max_workers=6) as ex:
 for result in ex.map(get,urls.items()):print(result)
