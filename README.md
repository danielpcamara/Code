# Sobre
Este repositório realiza o scraping de dados do site [econtabil](https://www.econtabil.com.br/) e armazena as informações em um banco de dados local sqlite. Nem todos os dados foram configurados para serem raspados. Os dados extraídos serão os seguintes:

* Item raspados com acesso admin:
    * 'Acesso ao menu... > Cadastro > Clientes' ```get_clients```
    * 'Acesso ao menu... > Cadastro > Clientes > Usuários' ```get_clients```
    * 'Acesso ao menu... > Cadastro > Usuários / Senhas' ```get_user```
    * 'Acesso ao menu... > Movimento > Impostos' ```get_mov```
* Itens raspados com permisão padrão
    * 'Acesso ao menu... > Cadastro > Solicitações' ```get_solicitacoes```
    * 'Acesso ao menu... > Cadastro > Outros Pagamentos' ```get_outros_pag```
    * 'Acesso ao menu... > Cadastro > Movimento Folha' (Eventos) ```get_movfolha```

# Objetivo
O foco destes scripts é baixar todos os Protocolos e Guias de impostos presentes no site de forma a criar um Backup dos dados localmente. Na sua configuração atual o script baixa em torno de 500 Guias e 500 Protocolos em 1 hora.

# Uso
Para verificar o seu uso recomenda-se observar o arquivo ```example.ipynb```.