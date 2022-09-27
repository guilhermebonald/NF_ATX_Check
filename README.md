Script de automação de processo para verificação de notas lançadas no software ATX Frota.

- Esse Script foi criado, devido ao esforço diario, para verificar notas fiscais lançadas no sistema fiscal ATX Frota.
- Foi usado as seguintes bibliotecas para esse mini projeto: 

OpenPyXl => Usado para coletar e inserir em uma lista uma série de números de notas fiscais.
PyAutogui => Todo o Script foi criado tendo como base central a biblioteca PyAutogui, para automatização dos cliques necessarios, e captura de tela.
Pytesseract => Usado para converter imagem em texto, e fazer a verificação pelo número da nota, se a mesma foi lançada no sistema.
Pandas => Geração de DataFrame e posteriormente convertendo em arquivo Excel. XLSX.
Win32Print e Win32Api => Utilizada para finalizando a geração do arquivo fazer e impressão do arquivo na impressora.
