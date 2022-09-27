<b>Script de verificação das notas fiscais lançadas no software ATX Frota.</b>

- Esse Script foi criado, devido ao trabalho repetitivo e diario, para verificar notas fiscais lançadas no sistema fiscal ATX Frota.
- Foi usado as seguintes bibliotecas para esse mini projeto: 

<b>OpenPyXl:</b> Usado para coletar e inserir em uma lista uma série de números de notas fiscais.

<b>PyAutogui:</b> Todo o Script foi criado tendo como base central a biblioteca PyAutogui, para automatização dos cliques necessarios, e captura de tela.

<b>Pytesseract:</b> Usado para converter imagem em texto, e fazer a verificação pelo número da nota, se a mesma foi lançada no sistema.

<b>Pandas:</b> Geração de DataFrame e posteriormente convertendo em arquivo Excel. XLSX.

<b>Win32Print e Win32Api:</b> Utilizada para finalizando a geração do arquivo fazer e impressão do arquivo na impressora.
