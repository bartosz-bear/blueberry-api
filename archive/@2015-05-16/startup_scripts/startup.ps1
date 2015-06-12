cd 'C:\Users\chbapie\AppData\Local\Continuum\Anaconda32\Scripts'
./activate apiquitous
cd "C:\Users\chbapie\Desktop\Bartosz\apiquitous"
Invoke-Command -ScriptBlock {.\python.exe api_server.py} -AsJob -ComputerName .