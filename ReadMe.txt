Este script Python tem como objetivo automatizar a atualização e o salvamento de uma planilha Excel. Ele realiza as seguintes tarefas:

- Abre a planilha: Carrega o arquivo Excel especificado no caminho definido pela variável caminho_planilha.
- Atualiza as consultas: Executa a atualização de todas as consultas presentes na planilha utilizando a função RefreshAll().
- Aguarda a conclusão das atualizações: Garante que todas as consultas sejam processadas antes de prosseguir.
- Salva a planilha: Salva as alterações realizadas na planilha.
- Fecha a planilha: Encerra a instância do Excel.


FUNCIONAMENTO
- Bibliotecas: Utiliza as bibliotecas time para controlar pausas na execução e win32com.client para interagir com o Microsoft Excel.
- Automatização: Realiza todas as ações de forma automatizada, sem a necessidade de interação manual do usuário.
- Visibilidade: O Excel é configurado para não exibir a interface gráfica durante a execução, tornando o processo mais discreto.

PERSONALIZAÇÃO

- Caminho da planilha: Modifique a variável caminho_planilha para indicar o local exato do arquivo Excel a ser atualizado.
- Tempo de espera: Os comandos time.sleep() podem ser ajustados para controlar o tempo de espera entre as diferentes etapas, caso necessário.