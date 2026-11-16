from fetch import get
import concurrent.futures
urls={'S5':'https://rocm.docs.amd.com/en/latest/conceptual/mi300.html','S6':'https://instinct.docs.amd.com/projects/system-acceptance/en/latest/gpus/mi300x.html','S7':'https://ualinkconsortium.org/specification/','S8':'https://www.businesswire.com/news/home/20260407620696/en/Ultra-Accelerator-Link-UALink-Consortium-Publishes-Four-Specifications-Defining-In-Network-Compute-Chiplets-Manageability-and-200G-Performance'}
with concurrent.futures.ThreadPoolExecutor(max_workers=4) as ex:
 for result in ex.map(get,urls.items()):print(result)
