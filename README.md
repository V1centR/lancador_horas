## Usage

### Java 11 jar application

Para executar a aplicação:

    $ java -jar geradorXlsHoras.jar <target year>


Example and result:

    $ java -jar geradorXlsHoras.jar 2022
    message: Arquivo Excel criado com sucesso!
    File output in: ./output/planilha_horas_2022.xls

## Observações

  - Caso ocorra o erro de "FileOutputStream" criar uma pasta com o nome "output" no mesmo diretorio da aplicação geradorXlsHoras.jar
