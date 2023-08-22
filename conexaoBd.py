import sqlite3 as bb

# teste de conexão com o banco de dados 

try:

    # variável para conectar a pasta do banco
    conn = bb.connect('bd/banco.db')

    # interage com o banco
    c = conn.cursor()

    c.execute('SELECT sqlite_version();') # executa a query

    version = c.fetchone() # guarda o resultado da query

    if version:
        print('Conexão realizada com sucesso !')
    else:
        print('Falha na conexão')

except bb.Error as erro:

    print('Falha na conexão: {erro}') # guarda na variável o erro entrado

finally:

    # fecha a conexão
    conn.close()
