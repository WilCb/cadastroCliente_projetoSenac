import PySimpleGUI as sg
import sqlite3 as bb
import aspose.words as aw

doc = aw.Document('relatorio.html')
doc.save('relatorio.pdf')

conn = bb.connect('bd/banco.db')
c = conn.cursor()

# Layout

layout = [
    [sg.Image('image/fundo.png', expand_x=True, expand_y=True)],
    [sg.Menu([
        ['Cadastro', ['Cadastro clientes', 'Cadastro Fornecedores', 'Cadastro Transportadoras']],
        ['Consulta', ['Consulta clientes', 'Consulta Fornecedores', 'Consulta Transportadoras']],
        ['Relatório', ['Relatório clientes', 'Relatório Fornecedores', 'Relatório Transportadoras']]
    ], tearoff=False)]
]

# criar a janela principal

janela = sg.Window('Sistema de vendas 1.0', layout, size=(400, 200), resizable=True)

while True:
    event, values = janela.read()

    if event == sg.WINDOW_CLOSED:
        break

    # acessa o menu cadastro -> cadastro clientes

    if event == 'Cadastro clientes':
        layoutCadastroClientes = [
            [sg.Text('Nome: '), sg.InputText(key='nome')],
            [sg.Text('E-mail: '), sg.InputText(key='email')],
            [sg.Text('Telefone:'), sg.InputText(key='telefone')],
            [sg.Button('Cadastrar'), sg.Button('Cancelar')]
        ]

        janelaCadastroClientes = sg.Window('Cadastro de clientes', layoutCadastroClientes, size=(400, 150))

        while True:
            event, values = janelaCadastroClientes.read()

            if event == sg.WINDOW_CLOSED or event == 'Cancelar':
                janelaCadastroClientes.close()
                break

            if event == 'Cadastrar':
                # Operações no Banco de Dados
                c.execute('INSERT INTO clientes (nome, email, telefone) VALUES(?, ?, ?)',
                (values['nome'], values['email'], values['telefone']))
                conn.commit()

                # Limpa os campos após efetuar o cadastro
                janelaCadastroClientes['nome'].update('')
                janelaCadastroClientes['email'].update('')
                janelaCadastroClientes['telefone'].update('')

                # Mostrar pop-up após o cadastro
                sg.popup('Cadastro efetuado', title='Cadastrado')
        janelaCadastroClientes.close()

    # acessa o menu Consulta -> Consulta clientes
    elif event == 'Consulta clientes':

        # Atualiza o registro alterado
        def edit_record(new_name, old_name):
            c.execute("UPDATE clientes SET nome=? WHERE nome=?", (new_name, old_name))
            conn.commit()


        # Deleta o registro escolhido
        def delete_record(name_to_delete):
            c.execute("DELETE FROM clientes WHERE nome=?", (name_to_delete,))
            conn.commit()


        # cria a tela de consulta
        consultaClientes = [
            [sg.Text('Nome do cliente: '), sg.InputText(key='nome')],
            [sg.Button('Todos'), sg.Button('Consultar por nome'), sg.Button('Cancelar')],
            [sg.Table(values=[], headings=['Nome', 'E-mail', 'Telefone'], expand_x=True, justification='c', display_row_numbers=False,
            auto_size_columns=False, num_rows=10, key='tabela')],
            [sg.Button('Novo Cliente'), sg.Button('Editar'), sg.Button('Excluir')]
        ]

        # variável para criar janela de consulta
        janelaConsultaClientes = sg.Window('Tela de Consulta de Clientes', consultaClientes, size=(600, 400), resizable=True)

        # loop infinito para manter a janela aberta
        while True:
            event, values = janelaConsultaClientes.read()

            if event == sg.WINDOW_CLOSED or event == 'Cancelar':
                janelaConsultaClientes.close()
                break

            if event == 'Consultar por nome':
                clienteBusca = values['nome'].upper().strip()
                c.execute('SELECT nome, email, telefone FROM clientes WHERE UPPER(nome) = ?', (clienteBusca,))

                registros = c.fetchall()

                # atualizar
                janelaConsultaClientes['tabela'].update(values=registros)
            
            elif event == 'Todos':
                c.execute('SELECT nome, email, telefone FROM clientes')
                registros = c.fetchall()
            
                janelaConsultaClientes['tabela'].update(values=registros)
            
            # cadastra um novo cliente dentro da tela de consulta
            elif event == 'Novo Cliente':
                layoutCadastroClientes = [
                    [sg.Text('Nome: '), sg.InputText(key='nome')],
                    [sg.Text('E-mail: '), sg.InputText(key='email')],
                    [sg.Text('Telefone:'), sg.InputText(key='telefone')],
                    [sg.Button('Cadastrar'), sg.Button('Fechar')]
                ]

                janelaCadastroClientes = sg.Window('Cadastro de clientes', layoutCadastroClientes, size=(400, 150))

                while True:
                    event, values = janelaCadastroClientes.read()

                    if event == sg.WINDOW_CLOSED or event == 'Fechar':
                        janelaCadastroClientes.close()
                        break

                    if event == 'Cadastrar':
                        c.execute('INSERT INTO clientes(nome, email, telefone) VALUES(?, ?, ?)', (values['nome'], values['email'], values['telefone']))
                        conn.commit()

                        janelaCadastroClientes['nome'].update('')
                        janelaCadastroClientes['email'].update('')
                        janelaCadastroClientes['telefone'].update('')

                        sg.popup('Cadastrado', title='Confirmação de cadastro')
                
            elif event == 'Editar':
                selected_row = values['tabela']
                if selected_row:
                    selected_row_index = selected_row[0]
                    row_data = registros[selected_row_index]
                    edited_name = sg.popup_get_text("Editar nome:", default_text=row_data[0])
                    if edited_name is not None:
                        old_name = row_data[0]
                        edit_record(edited_name, old_name)
                        registros[selected_row_index] = (edited_name, row_data[1])
                        janelaConsultaClientes['tabela'].update(values=registros)
                        sg.popup('Dados alterados !', title='Confirmação')

            elif event == 'Excluir':
                selected_row = values['tabela']
                if selected_row:
                    selected_row_index = selected_row[0]
                    row_data = registros[selected_row_index]
                    if sg.popup_yes_no('Tem certeza que deseja excluir este registro?', title='Confirmação') == 'Yes':
                        name_to_delete = row_data[0]
                        delete_record(name_to_delete)
                        registros.pop(selected_row_index)
                        janelaConsultaClientes['tabela'].update(values=registros)
                        sg.popup('Excluido !', title='Confirmação')

            # acessa o menu cadastro -> cadastro Fornecedores
    if event == "Cadastro Fornecedores":
        layoutCadastroFornecedores = [
            [sg.Text('CNPJ: '), sg.InputText(key='cnpj')],
            [sg.Text('Nome: '), sg.InputText(key='nome')],
            [sg.Button('Cadastrar'), sg.Button('Cancelar')],
        ]

        janelaCadastroFornecedor = sg.Window('Tela de cadastro fornecedor', layoutCadastroFornecedores, size=(400, 400))

        while True:
            event, values = janelaCadastroFornecedor.read()

            if event == sg.WINDOW_CLOSED or event == 'Cancelar':
                janelaCadastroFornecedor.close()
                break

            c.execute('INSERT INTO fornecedores(CNPJ, nome) VALUES(?, ?)', (values['cnpj'], values['nome']))
            conn.commit()

            janelaCadastroFornecedor['cnpj'].update('')
            janelaCadastroFornecedor['nome'].update('')

            sg.popup('Cadastro efetuado', title='Cadastro')

    # acessa o menu Consulta -> Consulta Fornecedores
    elif event == 'Consulta Fornecedores':

        # editar registro escolhido
        def edit_record(new_name, new_cnpj, old_name):
            c.execute("UPDATE fornecedores SET nome=?, CNPJ=? WHERE nome=?, CNPJ=?", (new_name, new_cnpj, old_name))
            conn.commit()


        # Deleta o registro escolhido
        def delete_record(name_to_delete):
            c.execute("DELETE FROM fornecedores WHERE nome=?", (name_to_delete,))
            conn.commit()


        # cria a tela de consulta
        layoutConsultaFornecedores = [
            [sg.Text('Nome do Fornecedor: '), sg.InputText(key='nome')],
            [sg.Button('Mostrar todos') ,sg.Button('Consultar'), sg.Button('Cancelar')],
            [sg.Table(values=[], headings=['Nome', 'CNPJ'], size=(30, 30), display_row_numbers=False,
            auto_size_columns=False, num_rows=10, key='tabela')],
            [sg.Button('Novo Fornecedor') ,sg.Button('Editar'), sg.Button('Excluir')]
        ]

        # variável para criar janela de consulta
        janelaConsultaFornecedores = sg.Window('Tela de Consulta de Fornecedores', layoutConsultaFornecedores, size=(400, 400))

        # loop infinito para manter a janela aberta
        while True:
            event, values = janelaConsultaFornecedores.read()

            if event == sg.WINDOW_CLOSED or event == 'Cancelar':
                janelaConsultaFornecedores.close()
                break

            # cria uma janela de cadastro dentro da seção de consulta
            if event == 'Novo Fornecedor':
                layoutCadastroFornecedores = [
                    [sg.Text('CNPJ: '), sg.InputText(key='cnpj')],
                    [sg.Text('Nome: '), sg.InputText(key='nome')],
                    [sg.Button('Cadastrar'), sg.Button('Cancelar')],
                ]

                janelaCadastroFornecedor = sg.Window('Tela de cadastro fornecedor', layoutCadastroFornecedores, size=(400, 400))

                while True:
                    event, values = janelaCadastroFornecedor.read()

                    if event == sg.WINDOW_CLOSED or event == 'Cancelar':
                        janelaCadastroFornecedor.close()
                        break

                    c.execute('INSERT INTO fornecedores(CNPJ, nome) VALUES(?, ?)', (values['cnpj'], values['nome']))
                    conn.commit()

                    janelaCadastroFornecedor['cnpj'].update('')
                    janelaCadastroFornecedor['nome'].update('')

                    sg.popup('Cadastro efetuado', title='Cadastro')

            elif event == 'Consultar':
                fornecedorBusca = values['nome'].upper().strip()
                c.execute('SELECT nome, CNPJ FROM fornecedores WHERE UPPER(nome) = ?', (fornecedorBusca,))

                registros = c.fetchall()

                # atualizar

                janelaConsultaFornecedores['tabela'].update(values=registros)

            elif event == 'Mostrar todos':
                c.execute('SELECT * FROM fornecedores')

                registros = c.fetchall()

                # atualizar

                janelaConsultaFornecedores['tabela'].update(values=registros)

           

    if event == 'Cadastro Transportadoras':

        layoutCadastroTransportadoras = [
            [sg.Text('Empresa:')],
            [sg.InputText(key='empresa')],
            [sg.Text('E-mail: ')],
            [sg.InputText(key='email')],
            [sg.Text('Telefone: ')],
            [sg.InputText(key='tel')],
            [sg.Text('Código de rastreamento: ')],
            [sg.InputText(key='cod')],
            [sg.Button('Cadastrar'), sg.Button('Cancelar')],
        ]

        janelaCadastroTransportadoras = sg.Window('Cadastro de Transportadoras', layoutCadastroTransportadoras, resizable=True)
        # loop para manter tela de cadastro aberta
        while True:
            event, values = janelaCadastroTransportadoras.read()

            if event == sg.WINDOW_CLOSED or event == 'Cancelar':
                janelaCadastroTransportadoras.close()
                break
            if event == 'Cadastrar':
                c.execute('INSERT INTO transportadora(empresa, email, telefone, cod_ratreio) VALUES(?, ?, ?, ?)', (values['empresa'], values['email'], values['tel'], values['cod']))
                conn.commit()

            # atualiza os inputs após cadastrar
            janelaCadastroTransportadoras['empresa'].update('')
            janelaCadastroTransportadoras['email'].update('')
            janelaCadastroTransportadoras['tel'].update('')
            janelaCadastroTransportadoras['cod'].update('')

            sg.popup('Cadastro efetuado', title='Cadastrado', auto_close=True)

    # CONSULTA DE TRANSPORTADORAS
    if event == 'Consulta Transportadoras':
            
            layoutConsultaTransportadoras = [
                [sg.Text('Nome da empresa: '), sg.InputText(key='empresa')],
                [sg.Button('Consultar'), sg.Button('Mostrar todos'), sg.Button('Cancelar')],
                [sg.Table(values=[], headings=('Empresa', 'E-mail', 'Telefone', 'Codigo de rastreio'), justification='c', expand_x=True, key='tabela', display_row_numbers=False,
            auto_size_columns=False, num_rows=10)],
                [sg.Button('Nova transportadora'), sg.Button('Editar'), sg.Button('Excluir')]
            ]

            janelaConsultaTransportadoras = sg.Window('Consultar transportadora', layoutConsultaTransportadoras, auto_size_buttons=True,resizable=True)

            while True:
                event, values = janelaConsultaTransportadoras.read()
                # fechar janela
                if event == sg.WINDOW_CLOSED or event == 'Cancelar':
                    janelaConsultaTransportadoras.close()
                    break
                
                if event == 'Nova transportadora':
                    layoutCadastroTransportadoras = [
                        [sg.Text('Empresa:')],
                        [sg.InputText(key='empresa')],
                        [sg.Text('E-mail: ')],
                        [sg.InputText(key='email')],
                        [sg.Text('Telefone: ')],
                        [sg.InputText(key='tel')],
                        [sg.Text('Código de rastreamento: ')],
                        [sg.InputText(key='cod')],
                        [sg.Button('Cadastrar'), sg.Button('Cancelar')],
                    ]

                    janelaCadastroTransportadoras = sg.Window('Cadastro de Transportadoras', layoutCadastroTransportadoras, resizable=True)
                    # loop para manter tela de cadastro aberta
                    while True:
                        event, values = janelaCadastroTransportadoras.read()

                        if event == sg.WINDOW_CLOSED or event == 'Cancelar':
                            janelaCadastroTransportadoras.close()
                            break
                        if event == 'Cadastrar':
                            c.execute('INSERT INTO transportadora(empresa, email, telefone, cod_ratreio) VALUES(?, ?, ?, ?)', (values['empresa'], values['email'], values['tel'], values['cod']))
                            conn.commit()

                        # atualiza os inputs após cadastrar
                        janelaCadastroTransportadoras['empresa'].update('')
                        janelaCadastroTransportadoras['email'].update('')
                        janelaCadastroTransportadoras['tel'].update('')
                        janelaCadastroTransportadoras['cod'].update('')

                        sg.popup('Cadastro efetuado', title='Cadastrado', auto_close=True)

                elif event == 'Consultar':

                    transportadoraBusca = values['empresa'].upper().strip()
                    c.execute('SELECT empresa, email, telefone, cod_ratreio FROM transportadora WHERE UPPER(empresa) = ?', (transportadoraBusca,))

                    registros = c.fetchall()

                    # atualizar
                    janelaConsultaTransportadoras['tabela'].update(values=registros)

                elif event == 'Mostrar todos':
                    c.execute('SELECT empresa, email, telefone, cod_ratreio FROM transportadora')

                    registros = c.fetchall()
                    janelaConsultaTransportadoras['tabela'].update(values=registros)
               


            

    ###################### RELATORIO ######################

    if event == 'Relatório clientes':

        # cria layout da janela de relatorio
        layoutTexto = [
            [sg.Text('Deseja gerar relátorio de clientes?'), sg.InputText(key='nome', visible=False)]
        ]

        layoutButton = [
            [sg.Button('Sim'), sg.Button('Não')]
        ]

        layoutRelatorioCentralizado = [
            [sg.Column(layoutTexto, justification='c')],
            [sg.Column(layoutButton, justification='c')]
        ]

        janelaRelatorio = sg.Window('Relatório de clientes', layoutRelatorioCentralizado, size=(400, 100))

        while True:
            event, values = janelaRelatorio.read()

            if event == sg.WINDOW_CLOSED or event == 'Não':
                janelaRelatorio.close()
                break

            clienteBuscar = values['nome'].upper()
            c.execute('SELECT * FROM clientes')
            registro = c.fetchall()

            if registro:
                with open('relatorio.html', 'w') as f:
                    f.write("<html><head><style> * {"
                            "margin: 0;"
                            "padding: 0;"
                            "}"
                            "body {"
                            "background-color: #ADD8E6;}"
                            "h1 {"
                            "text-align: center;"
                            "}"
                            ".container {"
                            "display: flex;"
                            "justify-content: center;"
                            "flex-direction: column;"
                            "}"
                            "td {"
                            "padding: 5px;"
                            "}"
                            "</style></head><body>")
                    f.write("<section class='container'><h1>Relatório</h1><table border='1'><tr><th>Nome</th><th>E-mail</th><th>Telefone</th></tr>")

                    for row in registro:
                        f.write(f'<tr><td>{row[0]}</td><td>{row[1]}</td><td>{row[2]}</td></tr>')
                    f.write('</table></section></body></html>')

                sg.popup('Relatório gerado com sucesso!', title='Relatório de Clientes')
            else:
                sg.popup('Produto não encontrado no banco de dados!', title='Relatório de Clientes')

            clienteBuscar = values['nome'].upper()
            c.execute('SELECT * FROM clientes WHERE UPPER(nome) = ?', (clienteBuscar,))
            registro = c.fetchall()

janela.close()

conn.close()