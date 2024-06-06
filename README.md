Tabela de Calendário Dinamica
Este script Python cria uma tabela de calendário resumida e intuitiva em formato Excel. Ele permite ao usuário definir uma data mínima e utiliza a data atual como data máxima, evitando a acumulação de datas fixas. Esta abordagem é particularmente útil para evitar o gerenciamento manual de datas em consultas do Power Query.

Funcionalidades
Geração Automática de Datas: O script utiliza a biblioteca Pandas para gerar automaticamente as datas no intervalo desejado, iniciando a partir de uma data mínima especificada pelo usuário e terminando na data atual.

Informações Adicionais sobre as Datas: Além das datas, o script adiciona informações úteis como o nome do mês, o número do mês e o ano correspondente.

Exportação para Excel: As informações são exportadas para um arquivo Excel (.xlsx), facilitando a visualização e o compartilhamento dos dados.

Como Usar
Configuração das Datas: No código-fonte, defina a data inicial desejada na variável data_inicial.

Execução do Script: Execute o script Python. Uma tabela de calendário será gerada e salva automaticamente em formato Excel no diretório especificado.

Pré-requisitos
Python 3.x
Bibliotecas Python: pandas, openpyxl
