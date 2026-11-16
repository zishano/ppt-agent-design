from pathlib import Path
exec((Path(__file__).resolve().parent / 'fetch.py').read_text().split('with concurrent.futures')[0])
urls={'S3':'https://www.nvidia.com/en-us/networking/products/infiniband/','S4':'https://www.nvidia.com/en-us/networking/products/ethernet/spectrum-x/','S5':'https://docs.nvidia.com/nemo/megatron-bridge/latest/parallelisms.html','S7':'https://docs.nvidia.com/deeplearning/nccl/user-guide/docs/usage/p2p.html'}
with concurrent.futures.ThreadPoolExecutor(max_workers=4) as ex:
 for result in ex.map(get,urls.items()):print(result)
