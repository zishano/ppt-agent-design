from fetch import get
import concurrent.futures,requests
from pathlib import Path
import pymupdf as fitz
R=Path(__file__).resolve().parent
urls={'S5':'https://ultraethernet.org/uec-1-0-whitepaper','S7':'https://docs.nvidia.com/networking/display/mlnxenv24102180lts/roce+congestion+control','S8':'https://www.ieee802.org/3/','S9':'https://www.ieee802.org/1/pages/802.1bb.html'}
with concurrent.futures.ThreadPoolExecutor(max_workers=4) as ex:
 for result in ex.map(get,urls.items()):print(result)
u='https://ultraethernet.org/wp-content/uploads/sites/20/2026/10/20261026-3711-UE-Specification-1.1-release-notes.pdf';a=requests.get(u,timeout=30);a.raise_for_status();(R/'S10.pdf').write_bytes(a.content);d=fitz.open(stream=a.content,filetype='pdf');t='\n'.join(f'PAGE {i+1}\n'+p.get_text() for i,p in enumerate(d));(R/'S10.txt').write_text(t);(R/'S10.url').write_text(u);print(t[:11000])
