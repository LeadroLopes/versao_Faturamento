from tkinter import *
from tkinter import ttk
from tkinter import messagebox
from tkcalendar import Calendar, DateEntry
from babel.numbers import *
import pyodbc

app = Tk()


class Funcoes:
    def __init__(self):
        self.app = app
        self.cor_treeview()
        # self.sql()
        self.tela()
        self.frame()
        self.frame1()
        # self.frame2()
        self.treeview_app()
        # self.desconectar()
        app.mainloop()
#                                                   FUNÇOES DA TELA PAI                                            #

    def tela(self):
# ESTA A BARRA DE MENUS  EAS CONFIGURAÇOES DA TELA APP #

        self.app.configure(bg='#421212')
        # self.app.state('zoomed')
        self.app.resizable(False, False)
        self.app.geometry('1350x700+01+3')
        # self.app.overrideredirect(True)
        self.barraMenu = Menu(self.app)
        self.consulta_exporta = Menu(self.barraMenu, tearoff=0)
        self.consulta_exporta.add_command(label='CONSULTAR E EXPORTA ESTOQUE', command=self.exportar_estoque)
        self.consulta_exporta.add_separator()
        self.consulta_exporta.add_command(label='CONSULTA VENDA/CLIENTE', command=self.exporta_vendas)
        self.consulta_exporta.add_separator()
        self.consulta_exporta.add_command(label='REPOSIÇAO DE ESTOQUE', command=self.reposicao)
        self.consulta_exporta.add_separator()
        self.consulta_exporta.add_command(label='PLANILHA PARA ATACADO', command=self.planila)
        self.consulta_exporta.add_separator()
        self.consulta_exporta.add_command(label='FATURAMENTO', command=self.faturamento)
        self.consulta_exporta.add_separator()
        self.consulta_exporta.add_command(label='FECHAR', command=self.app.quit)
        self.consulta_exporta.add_separator()
        self.barraMenu.add_cascade(label='OPÇÕES', menu=self.consulta_exporta)

        self.app.config(menu=self.barraMenu)  # ESTA A BARRA DE MENUS  EAS CONFIGURAÇOES DA TELA APP ####

# CONFIGURANDO O TREEVIEW

    def cor_treeview(self):
        style = ttk.Style()
        style.theme_use("default")
        style.configure("Treeview",

                        foreground="#000000",
                        rowheight=20,
                        fieLdbackground="silver"
                        )
        # COR DA SELEÇAO DO TRIVIEW

        style.map("Treeview",
                  background=[('selected', 'green')])
    #                                    NA FUNÇAO FRAME ESTA OS BOTES JUNTO A FRaME BOTOES                           #

    def frame(self):  # botoes
        self.frame = Frame(width=380, height=650, borderwidth=9, relief="raised",
                           highlightthickness=2, background='#871003').place(x=10, y=10)

        self.img_produto = PhotoImage(file='imagens/produto.png')
        self.bt_cadastrar = Button(text='CADRASTRAR', command=self.cadastrar, image=self.img_produto, compound=LEFT,
                                   background='#321a01',
                                   foreground='#ffffff',
                                   padx=20, pady=10,
                                   font='times 15',
                                   borderwidth=5, relief="raised")
        self.bt_cadastrar.place(x=50, y=30)

        self.carrinho = PhotoImage(file='imagens/carrinho1.png')
        self.bt_realizar_venda = Button(text='REALIZAR VENDAS', image=self.carrinho, compound=LEFT,
                                        background='#321a01',
                                        foreground='#ffffff',
                                        padx=14, pady=5,
                                        font='times 13',
                                        borderwidth=5, relief="raised", command=self.vendas)
        self.bt_realizar_venda.place(x=50, y=140)
        self.alterar = PhotoImage(file='imagens/alterar1.png')
        self.bt_alterar_valor = Button(text='ALTERAR VALORES', image=self.alterar, compound=LEFT,
                                       background='#321a01',
                                       foreground='#ffffff',
                                       padx=2,
                                       font='times 15',
                                       borderwidth=5, relief="raised", command=self.alterar_valores)
        self.bt_alterar_valor.place(x=50, y=240)
        self.deletar = PhotoImage(file='imagens/deletar1.png')

        self.bt_deletar = Button(text='DELETE', image=self.deletar, compound=LEFT,
                                 background='#a20f0f',
                                 foreground='white',
                                 padx=35,
                                 font='times 19',
                                 borderwidth=5, relief="raised", command=self.deletar_item_tabela)
        self.bt_deletar.place(x=50, y=555)
        self.consulta = PhotoImage(file='imagens/consultar_vendas.png')

        self.bt_consultar_venda = Button(text='CONSULTAR VENDAS', image=self.consulta, compound=LEFT,
                                         background='#321a01',
                                         foreground='white',
                                         pady=15,
                                         font='times 15',
                                         borderwidth=5, relief="raised",
                                         command=self.consultar_vendas)
        self.bt_consultar_venda.place(x=50, y=340)
        self.configurar = PhotoImage(file='imagens/configura_estoque.png')
        self.bt_configurar_estoque = Button(text='CONFIGURAR ESTOQUE', image=self.configurar, compound=LEFT,
                                            background='#321a01',
                                            foreground='white',
                                            padx=7,
                                            pady=7,
                                            font='times 11',
                                            borderwidth=5, relief="raised",
                                            command=self.estoque_minimo, )
        self.bt_configurar_estoque.place(x=50, y=455)

    def frame1(self):
        self.frame = Frame(width=950, height=655, borderwidth=9, relief="raised",
                           highlightthickness=2, background='#871003',
                          )
        self.frame.place(x=400, y=10)
        self.label = Label(text='PESQUISAR POR SEGUIMENTO', font='times 30', bg='#871003', fg='#ffffff')
        self.label.place(x=560,y=30)

        self.label = Label(text='SEGUIMENTO', font='times 17 bold', bg='#871003', fg='white')
        self.label.place(x=660, y=120)

        self.label_tema = Label(text='TEMA', font='times 17 bold', bg='#871003', fg='white')
        self.label_tema.place(x=565, y=160)

        self.pesquisa_app = Entry(foreground='black', font='times 15 bold')
        self.pesquisa_app.place(x=410, y=120, width=245)

        self.pesquisa_tema_app = Entry(foreground='#ff6a06', font='times 15')
        self.pesquisa_tema_app.place(x=410, y=160, width=150)

        self.labell = Label(text='CODIGO DO SEGUIMENTO', font='times 17 bold', bg='#871003', fg='white')
        self.labell.place(x=980, y=100)

        self.entry_pesquisa1 = Entry(foreground='#ff6a06', font='times 15')
        self.entry_pesquisa1.place(x=1050, y=140, width=200)

        self.imgE= PhotoImage(file='imagens/pesquisadetalhada.png')
        self.busca = Button(text='PESQUISAR',image=self.imgE,
                            background='#871003',
                            foreground='black',
                            font='times  10 bold',
                            borderwidth=5, relief=FLAT, command=self.pesquisar_app)
        self.busca.place(x=650, y=153)

        self.total_seguimento=Button(self.app,text='SEGUIMENTOS TOTAIS',command=self.seguimento_total,bg='#c3bebe'
                                     ,relief=FLAT,font='times 12 bold')
        self.total_seguimento.place(x=410,y=210)

        self.total_quantidade = Button(self.app, text='QTDD TOTAL EM ESTOQUE',bg='#c3bebe',
                                       command=self.quantidade_total,relief=FLAT,font='times 12 bold')
        self.total_quantidade.place(x=700,y=210)

        self.total_valor = Button(self.app, text='VALOR TOTAL EM ESTOQUE ',font='times 12 bold',
                                  bg='#c3bebe', command=self.valor_total_em_estoque)
        self.total_valor.place(x=990,y=210)

    def treeview_app(self):
        self.columns = ('id', 'seguimento', 'tema', 'descricao', 'tamanho', 'quantidade', 'valor')
        self.treeview = ttk.Treeview(self.app, columns=self.columns, show='headings')

        self.treeview.heading('id', text='ID')
        self.treeview.heading('seguimento', text='Seguimento')
        self.treeview.heading('tema', text='Tema')
        self.treeview.heading('descricao', text='Descricao')
        self.treeview.heading('tamanho', text='Tamanho')
        self.treeview.heading('quantidade', text='Quantidade')
        self.treeview.heading('valor', text='Valor')

        self.treeview.column('id', width=80)
        self.treeview.column('seguimento', width=180)
        self.treeview.column('tema', width=180)
        self.treeview.column('descricao', width=180)
        self.treeview.column('tamanho', width=50)
        self.treeview.column('quantidade', width=50)
        self.treeview.column('valor', width=50)
        self.treeview.place(x=410, y=250, width=895, height=400)
        self.scrolll = Scrollbar(self.app, orient=VERTICAL)
        self.scrolll = Scrollbar(self.app, orient=VERTICAL, command=self.treeview.yview)
        self.treeview.configure(yscrollcommand=self.scrolll.set)
        self.scrolll.place(x=1310, y=250, width=25, height=400)
        # fim da teeview

    def sql(self):
        self.conectar = self.conexao_dados = ("Driver={SQL Server};"  # criando a conexao
                                              "Server=DESKTOP-RFPLHG0;"
                                              "Database=seguimentos_cadastrados;")

        self.conexao = pyodbc.connect(self.conectar)
        self.cursor = self.conexao.cursor()

    def desconectar(self):
        # self.conexao.close()
        self.cursor.close()

        print('desconectado')

    def pesquisar_app(self):

        self.treeview.delete(*self.treeview.get_children())
        self.sql()
        self.pesquisa_appp = self.pesquisa_app.get().strip()
        self.pesquisa_app_tema=self.pesquisa_tema_app.get().strip()
        self.pesquisa_app.delete(0, END)

        try:
            if self.pesquisa_appp != '' and self.pesquisa_app_tema != '':
                seguimento = self.cursor.execute(
                    f"""SELECT   * FROM seguimento_venda  WHERE seguimento like '{self.pesquisa_appp}%'
                                                               AND tema like '{self.pesquisa_app_tema}%'
                        ORDER BY seguimento,tema,tamanho DESC""")
                for linha in seguimento:

                    self.y = linha.id, linha.seguimento, linha.tema, linha.descricao, linha.tamanho, \
                             linha.quantidade, linha.valor
                    self.treeview.insert('', 'end', values=self.y)
            elif self.pesquisa_app_tema != '':
                tema= self.cursor.execute(
                    f"""SELECT TOP 50 * FROM seguimento_venda  WHERE tema like '{self.pesquisa_app_tema}%'
                        ORDER BY seguimento,tema,tamanho DESC""")
                for linha in tema:
                    self.t = linha.id, linha.seguimento, linha.tema, linha.descricao, linha.tamanho, \
                             linha.quantidade, linha.valor
                    self.treeview.insert('', 'end', values=self.t)

            elif self.pesquisa_appp != '':
                seguimento = self.cursor.execute(
                    f"""SELECT * FROM seguimento_venda  WHERE seguimento like '{self.pesquisa_appp}%'
                            ORDER BY seguimento,tema,tamanho DESC""")
                for linha in seguimento:
                    self.s = linha.id, linha.seguimento, linha.tema, linha.descricao, linha.tamanho, \
                             linha.quantidade, linha.valor
                    self.treeview.insert('', 'end', values=self.s)

                # if self.pesquisa_appp not in seguimento:
                #     messagebox.showinfo('','nao cadastrado')
            else:
                self.pesquisa1 = self.cursor.execute(f"""
                                        SELECT DISTINCT top 100 seguimento,tema,valor
                                        FROM seguimento_venda  WHERE seguimento IS NOT NULL""")
                for linha in self.cursor.fetchall():
                    self.linha = linha.seguimento, linha.tema, linha.valor
                    self.treeview.insert('', 'end', values=self.linha)
            self.desconectar()
        except:
            pass

    def seguimento_total(self):
        self.sql()
        qtdd_estoque=self.cursor.execute("""SELECT COUNT(seguimento) from seguimento_venda""")
        for s in qtdd_estoque:
            total=s[0]
            if total != '':
                messagebox.showinfo('INFORMAÇAO',f'VOCÊ ESTA COM  {total:.2f} SEGUIMENTOS CADASTRO EM SEU ESTOQUE')

    def quantidade_total(self):
        self.sql()
        qtdd_estoque=self.cursor.execute("""SELECT sum(quantidade) from seguimento_venda""")
        for s in qtdd_estoque:
            total=s[0]
            if total != '':
                messagebox.showinfo('',f'VOCÊ TEM  {total:,} PEÇAS EM SEU ESTOQUE NO MOMENTO ATUAL')

    def valor_total_em_estoque(self):
        self.sql()
        qtdd_estoque=self.cursor.execute("""SELECT sum(quantidade * valor) from seguimento_venda""")
        for s in qtdd_estoque:
            total=s[0]
            if total != '':
                messagebox.showinfo('',f'VOCÊ TEM O VALOR DE {total:,} EM ESTOQUE')


    #                          FUNÇOES DE CADASTRO  LIGADO AO BOTAO CADASTRAR                                          #


    def cadastrar(self):
        self.app2 = Toplevel()
        self.app2.geometry('930x580+333+56')
        self.app2.resizable(False, False)
        self.app2.transient(self.app)
        self.app2.focus_force()
        self.app2.grab_set()

        self.app2.configure(bg='#2a2a29')
        #                        AQUI SAO AS FRAMES SUPERIOR E INFERIOR                                        #

        self.frame_cadastro = Frame(self.app2, bg="white", width=915, height=290, borderwidth=9, relief="raised",
                                    background='#7f7f7f')
        self.frame_cadastro.place(x=10, y=10)

        self.frame_cadastro_inferior = Frame(self.app2, bg="white", width=915, height=260, borderwidth=9,
                                             relief="raised",
                                             background='#7f7f7f')
        self.frame_cadastro_inferior.place(x=10, y=310)
        self.lb_cadastro = Label(self.app2, text='ARÉA DE CADASTRO', bg='#574c26',pady=6,
                                 fg='white', font='times 15', relief='raised')
        self.lb_cadastro.place(x=350, y=18)

#                               AQUI SAO AS LABEL ORDENADA IGUAL AS COLUNAS DO BANCO DE DADOS                         #

        self.lb_cadastro_id = Label(self.app2, text='ID', bg='#7f7f7f',
                                    fg='white', relief=FLAT, font='times 14 bold')
        self.lb_cadastro_id.place(x=475, y=60)
        self.lb_cadastro_seguimento = Label(self.app2, text='SEGUIMENTO', bg='#7f7f7f',
                                            fg='white', font='times 14', relief='raised')
        self.lb_cadastro_seguimento.place(x=290, y=60)
        self.lb_cadastro_tema = Label(self.app2, text='TEMA/COR(SEG)', font='times 14', bg='#7f7f7f',
                                      fg='white', relief='raised')
        self.lb_cadastro_tema.place(x=290, y=100)
        self.lb_cadastro_descricao = Label(self.app2, text='DESCRIÇAO', bg='#7f7f7f',
                                           fg='white', relief='raised', font='times 14')
        self.lb_cadastro_descricao.place(x=290, y=140)
        self.lb_cadastro_tamanho = Label(self.app2, text='TAMANHO', bg='#7f7f7f',
                                         fg='white', relief='raised', font='times 14')
        self.lb_cadastro_tamanho.place(x=120, y=175)
        self.lb_cadastro_quantidade = Label(self.app2, text='QUANTIDADE', bg='#7f7f7f',
                                            fg='white', relief='raised', font='times 14')
        self.lb_cadastro_quantidade.place(x=120, y=212)
        self.lb_cadastro_valor = Label(self.app2, text='VALOR', bg='#7f7f7f',
                                       fg='white', relief='raised', font='times 14')
        self.lb_cadastro_valor.place(x=120, y=252)

#                       AQUI SAO MEUS ENTRYS TODOS ELES SEGUEM O MSM PADRAO DO BANCO DE DADOS                          #

        self.entry_cadastro_id = Entry(self.app2, font='times 15 bold ', width=5, fg='white', bg='#7f7f7f',relief=FLAT)
        self.entry_cadastro_id.place(x=420, y=60)

        self.entry_cadastro_seguimento = Entry(self.app2, font='times 15', width=25)
        self.entry_cadastro_seguimento.place(x=30, y=60)

        self.entry_cadastro_tema = Entry(self.app2, font='times 15', width=25)
        self.entry_cadastro_tema.place(x=30, y=100)

        self.entry_cadastro_descricao = Entry(self.app2, font='times 15', width=25)
        self.entry_cadastro_descricao.place(x=30, y=140)

        self.entry_cadastro_tamanho = Entry(self.app2, font='times 15', width=7)
        self.entry_cadastro_tamanho.place(x=30, y=175)

        self.entry_cadastro_quantidade = Entry(self.app2, font='times 15 ', width=7)
        self.entry_cadastro_quantidade.place(x=30, y=213)

        self.entry_cadastro_valor = Entry(self.app2, font='times 15 ', width=7)
        self.entry_cadastro_valor.place(x=30, y=253)

#               LABEL ESQUEDAR DA FUNCAO CADASTRAR ONDE REALIZO A BUSCA PARA CORRIGIR OS VALORES                       #

        self.lab_cadastro = Label(self.app2, text='BUSCA POR SEGUIMENTO', bg='#6b6a68', fg='white', relief='raised',
                                  font='times 15')
        self.lab_cadastro.place(x=550, y=70)

        self.lab_cadastro_tema = Label(self.app2, text='TEMA', bg='#6b6a68', fg='white', relief='raised',
                                       font='times 15')
        self.lab_cadastro_tema.place(x=640, y=135)
        self.entry_cadastro_pesquisa = Entry(self.app2, font='times 15', width=25)
        self.entry_cadastro_pesquisa.place(x=510, y=105)

        self.entry_cadastro_pes_tema = Entry(self.app2, font='times 15', width=12)
        self.entry_cadastro_pes_tema.place(x=510, y=135)
    #                            BOTAO QUE REALIZA A BUSCA PARA AÇTERAR VALOR                                       #

        self.cadastro_vendas = Button(self.app2, text='BUSCAR', bg='#574c26', fg='white', font='times 12',
                                      command=self.pesquisar_cadastro)
        self.cadastro_vendas.place(x=710, y=140)

        self.cadastro = PhotoImage(file='imagens/cadastro.png')
        self.bt_cadastro_pesquisar = Button(self.app2, text='CADASTRAR', bg='#7f7f7f', pady=5,image=self.cadastro,
                                            font='times',compound=TOP,fg='#f4f23d',relief=FLAT,
                                            command=self.inserir_dados_treeview_cadastrar)
        self.bt_cadastro_pesquisar.place(x=400, y=200)
        self.corrigir= PhotoImage(file='imagens/corrigir.png')
        self.bt_cadastro_corrigir = Button(self.app2, text=' CORRIGIR' , bg='#574c26', fg='white', pady=2,
                                           image=self.corrigir,compound=LEFT,
                                           font='times 13 bold', command=self.aterar_valores_cadastro)
        self.bt_cadastro_corrigir.place(x=25, y=17)

        self.bt_cadastro_voltar = Button(self.app2, text='VOLTAR', command=self.app2.destroy, bg='#7f7f7f', fg='#f4f23d',
                                         relief=FLAT, pady=5,
                                         font='times 15')
        self.bt_cadastro_voltar.place(x=790, y=20)

        self.treeview_cadastroc()

    def treeview_cadastroc(self):
        self.columns = ('id', 'seguimento', 'tema', 'descricao', 'tamanho', 'quantidade', 'valor')
        self.treeview_cadastro = ttk.Treeview(self.app2, columns=self.columns, show='headings')
        self.treeview_cadastro.heading('id', text='ID')
        self.treeview_cadastro.heading('seguimento', text='Seguimento')
        self.treeview_cadastro.heading('tema', text='Tema')
        self.treeview_cadastro.heading('tamanho', text='Tamanho')
        self.treeview_cadastro.heading('descricao', text='Descriçao')
        self.treeview_cadastro.heading('quantidade', text='Quantidade')
        self.treeview_cadastro.heading('valor', text='Valor')

        self.treeview_cadastro.column('id', width=40)
        self.treeview_cadastro.column('seguimento', width=200)
        self.treeview_cadastro.column('tema', width=180)
        self.treeview_cadastro.column('tamanho', width=50)
        self.treeview_cadastro.column('descricao', width=50)
        self.treeview_cadastro.column('quantidade', width=50)
        self.treeview_cadastro.column('valor', width=50)
        self.treeview_cadastro.place(x=20, y=320, width=870, height=240)
        self.scrollcad = Scrollbar(self.app2, orient=VERTICAL, command=self.treeview_cadastro.yview)
        self.treeview_cadastro.configure(yscrollcommand=self.scrollcad.set)
        self.scrollcad.place(x=895, y=320, width=25, height=240)
        self.treeview_cadastro.bind("<Double-1>", self.double_click_cadastro)
    #                FUNÇAO INSERIR,RESPONSAVEL POR INSERIR E RETORNAO SEGUIMENTOS DUPLICADO CASO AJA ALGUM            #

    def inserir_dados_treeview_cadastrar(self):
        self.sql()
        self.treeview_cadastro.delete(*self.treeview_cadastro.get_children())
        self.cadastrar_id = self.entry_cadastro_id.get().strip()
        self.cadastrar_segui = self.entry_cadastro_seguimento.get().strip().capitalize()
        self.cadastrar_tema = self.entry_cadastro_tema.get().strip().capitalize()
        self.cadastrar_descricao = self.entry_cadastro_descricao.get().strip().capitalize()
        self.cadastrar_tamanho = self.entry_cadastro_tamanho.get().strip().capitalize()
        self.cadastrar_quantidade = self.entry_cadastro_quantidade.get().strip()
        self.cadastrar_valor = self.entry_cadastro_valor.get().strip().replace(',', '.')
        self.teste1 = self.cadastrar_segui, self.cadastrar_tema, self.cadastrar_tamanho, self.cadastrar_descricao
        try:
            if self.cadastrar_segui == '' or \
                    self.cadastrar_tema == '' or \
                    self.cadastrar_descricao == '' or \
                    self.cadastrar_tamanho == '' or \
                    self.cadastrar_quantidade == '' or \
                    self.cadastrar_valor == '':
                messagebox.showinfo('PREENCHA!', 'PREECHA OS CAMPOS CORRETAMENTE PARA REALIZAR O CADASTRO')
            else:
                self.mensg = messagebox.askyesno('CADASTRO', f'DESEJA  CADASTRAR O SEGUIMENTO {self.cadastrar_segui}?')
                if self.mensg == TRUE:
                    self.cp = self.cursor.execute(
                        f"""INSERT INTO seguimento_venda(seguimento,tema,descricao,tamanho,quantidade,valor)
                                           VALUES
                                           ('{self.cadastrar_segui}',
                                           '{self.cadastrar_tema}',
                                           '{self.cadastrar_descricao}',
                                           '{self.cadastrar_tamanho}',
                                           '{self.cadastrar_quantidade}',
                                           {self.cadastrar_valor})
                                           """)

                    self.p = self.cursor.execute(f"""
                                SELECT id,seguimento,tema,descricao,tamanho,quantidade,valor
                                FROM seguimento_venda
                                WHERE seguimento = '{self.cadastrar_segui}' 
                                AND tema ='{self.cadastrar_tema}'
                                AND tamanho = '{self.cadastrar_tamanho}'
                                AND descricao= '{self.cadastrar_descricao}'ORDER BY tamanho DESC
                                """)
                    messagebox.showinfo('CONCLUIDO', 'SEGUIMENTO CADASTRADO COM SUCESSO')
                    for linha in self.cursor.fetchall():
                        self.v = linha.id,\
                                 linha.seguimento,\
                                 linha.tema,\
                                 linha.descricao,\
                                 linha.tamanho,\
                                 linha.quantidade,\
                                 linha.valor
                        self.treeview_cadastro.insert('', 'end', values=self.v)

                    self.mensg1 = messagebox.askyesno(
                        '', f'DESEJA CONTINUAR A CADASTRAR O SEGUIMENTO {self.cadastrar_segui}'
                    )
                    if self.mensg1 == TRUE:
                        # self.entry_cadastro_tema.delete(0, END)
                        # self.entry_cadastro_descricao.delete(0, END)
                        self.entry_cadastro_tamanho.delete(0, END)
                    else:
                        messagebox.showinfo(
                            '', 'OKAY, PARA REALIZAR UM NOVO CADASTRO E NECESSARIO PREENCHER TODOS OS CAMPOS NOVAMENTE'
                        )
                        self.limpar_tela_treeview_cadastro()

                    self.cursor.commit()
                    self.desconectar()
                else:
                    messagebox.showerror('', 'NENHUM SEGUIMENTO SERÁ CADASTRADO')
        except:
            messagebox.showerror('VALORES INVALIDOS', 'PREENCHA OS VALORES CORRETAMENTE')

    #               PESQUISAR CADASTRO LIGADO AO BOTAO BUSCAR DA FUNÇAO CADASTRAR  RESPONSAVEL PARA CORRIGIR VALORES   #

    def pesquisar_cadastro(self):
        self.treeview_cadastro.delete(*self.treeview_cadastro.get_children())
        self.sql()
        self.busca = self.entry_cadastro_pesquisa.get().strip()
        self.busca_tema = self.entry_cadastro_pes_tema.get().strip()

        if self.busca == '':
            messagebox.showinfo('PESQUISA DE SEGUIMENTO', 'DIGITE O SEGUIMENTO QUE DESEJA REALIZAR A BÚSCA')
        elif self.busca != '' and self.busca_tema != '':
            self.pesquisar1 = self.cursor.execute(f"""SELECT * from seguimento_venda
                                                  WHERE  seguimento like '{self.busca}%'
                                                    AND tema LIKE '{self.busca_tema}%' ORDER BY tamanho DESC """)

            for linha in self.cursor.fetchall():
                self.buscar1 = linha.id, \
                               linha.seguimento, \
                               linha.tema, \
                               linha.descricao, \
                               linha.tamanho, \
                               linha.quantidade, \
                               linha.valor
                self.treeview_cadastro.insert('', 'end', values=self.buscar1)
        else:
            self.pesquisar = self.cursor.execute(f"""SELECT * from seguimento_venda
                                                  WHERE  seguimento like '{self.busca}%'  ORDER BY tamanho DESC """)
            for linha in self.cursor.fetchall():
                self.buscar = linha.id, \
                              linha.seguimento, \
                              linha.tema, \
                              linha.descricao, \
                              linha.tamanho, \
                              linha.quantidade, \
                              linha.valor
                self.treeview_cadastro.insert('', 'end', values=self.buscar)

    def aterar_valores_cadastro(self):
        self.sql()
        self.id = self.entry_cadastro_id.get().strip()
        self.tamanho = self.entry_cadastro_tamanho.get().strip().capitalize()
        self.tema = self.entry_cadastro_tema.get().strip().capitalize()
        self.descri = self.entry_cadastro_descricao.get().strip().capitalize()
        self.comparar = self.entry_cadastro_seguimento.get().strip().capitalize()
        self.quantidade = self.entry_cadastro_quantidade.get().strip()
        try:
            if self.tamanho != '' or self.tema != '' or self.descri != '' or self.comparar != '' or self.quantidade != '' :
                aviso= messagebox.askyesno('','DESEJA FAZER A ALTERAÇOES DOS DADOS CADASTRADO?')
                if aviso == TRUE:
                    self.alterardados= self.cursor.execute(
                        f""" UPDATE seguimento_venda SET seguimento = '{self.comparar}',
                        tema = '{self.tema}',
                        descricao = '{self.descri}',
                        tamanho = '{self.tamanho}',
                        quantidade = '{self.quantidade}'
                        WHERE id = '{self.id}' """)
                    messagebox.showinfo('','ALTERAÇAO REALIZADA COM SUCESSO')
                    self.cursor.commit()
                    self.desconectar()
                else:
                    messagebox.showinfo('','NENHUM REGISTRO FOI ALTERADO FOI ALTERADA')
            else:
                messagebox.showinfo('','E NECESSARIO PREENCHER OS CAMPOS\n'
                                       'PARA REALIZAR A CORREÇÃO')
        except:
            messagebox.showerror('ERRO!',f'O QUATIDADE {self.quantidade} NAO E UMA QUANTIDADE VALIDA\n'
                                         ' POR FAVOR INFORME CORRETAMENTE A QUANTIDADE DESEJADA ')

    def limpar_tela_treeview_cadastro(self):
        self.entry_cadastro_id.delete(0, END)
        self.entry_cadastro_seguimento.delete(0, END)
        self.entry_cadastro_tema.delete(0, END)
        self.entry_cadastro_descricao.delete(0, END)
        self.entry_cadastro_tamanho.delete(0, END)
        self.entry_cadastro_quantidade.delete(0, END)
        self.entry_cadastro_valor.delete(0, END)

    def double_click_cadastro(self, event):
        self.limpar_tela_treeview_cadastro()
        self.treeview_cadastro.selection()
        for y in self.treeview_cadastro.selection():
            col1, col2, col3, col4, col5, col6, col7 = self.treeview_cadastro.item(y, 'values')
            self.entry_cadastro_id.insert(END, col1)
            self.entry_cadastro_seguimento.insert(END, col2)
            self.entry_cadastro_tema.insert(END, col3)
            self.entry_cadastro_descricao.insert(END, col4)
            self.entry_cadastro_tamanho.insert(END, col5)
            self.entry_cadastro_quantidade.insert(END, col6)
            self.entry_cadastro_valor.insert(END, col7)
            self.entry_cadastro_valor.delete(0, END)

 #                                             FUNÇAO REALIZAR VENDAS                                                 #

    def vendas(self):
        self.app5 = Toplevel()
        self.app5.geometry('970x600+290+56')
        self.app5.resizable(False, False)
        self.app5.transient(self.app)
        self.app5.focus_force()
        self.app5.grab_set()
        self.app5.configure(bg='#2a2a29')

        self.framevendas = Frame(self.app5, bg="white", width=955, height=300, borderwidth=9, relief="raised",
                                 background='#7f7f7f')
        self.framevendas.place(x=10, y=10)

        self.frameinferiorv = Frame(self.app5, bg="white", width=955, height=270, borderwidth=9, relief="raised",
                                    background='#7f7f7f')
        self.frameinferiorv.place(x=10, y=315)

        self.lb_vendas = Label(self.app5, text='REALIZAR VENDAS ', bg='#574c26',
                               fg='white', font='times 15', relief='raised')
        self.lb_vendas.place(x=350, y=23)

        self.lb_vendas_Produto = Label(self.app5, text='CLIENTE', bg='#7f7f7f',
                                       fg='white', relief='raised', font='times 14')
        self.lb_vendas_Produto.place(x=190, y=58)

        self.lb_vendas_SeguiMento = Label(self.app5, text='SEGUIMENTO', bg='#7f7f7f',
                                          fg='white', relief='raised', font='times 14')
        self.lb_vendas_SeguiMento.place(x=190, y=93)

        self.lb_vendas_TEMA = Label(self.app5, text='TEMA ou COR', bg='#7f7f7f',
                                    fg='white', relief='raised', font='times 14')
        self.lb_vendas_TEMA.place(x=190, y=123)

        self.lb_vendas_descricao = Label(self.app5, text='DESCRIÇAO', bg='#7f7f7f',
                                         fg='white', relief='raised', font='times 14')
        self.lb_vendas_descricao.place(x=190, y=165)
        self.lb_vendas_Quantidade = Label(self.app5, text='TAMANHO', bg='#7f7f7f',
                                          fg='white', relief='raised', font='times 14')
        self.lb_vendas_Quantidade.place(x=110, y=200)

        self.lb_vendas_VALOR = Label(self.app5, text='QUANTIDADE', bg='#7f7f7f',
                                     fg='white', relief='raised', font='times 14')
        self.lb_vendas_VALOR.place(x=110, y=235)

        self.lb_vendas_VALOR = Label(self.app5, text='VALOR', bg='#7f7f7f',
                                     fg='white', relief='raised', font='times 14')
        self.lb_vendas_VALOR.place(x=110, y=270)
        self.lb_teste = Label(self.app5, text='PESQUISAR  SEGUIMENTO', bg='#7f7f7f',
                              fg='white', relief='raised', font='times 14 bold')
        self.lb_teste.place(x=370, y=160)
        self.treeview_vendas_valores()

#                                  MEUS ENTRYS DA FUNÇAO                                                              #
        self.Entry_vendas_cliente = Entry(self.app5, font='times 15', width=15)
        self.Entry_vendas_cliente.place(x=25, y=58)
        self.Entry_vendas_SEGUIMENTO = Entry(self.app5, font='times 15', width=15)
        self.Entry_vendas_SEGUIMENTO.place(x=25, y=95)
        self.Entry_vendas_TEMA = Entry(self.app5, font='times 15', width=15)
        self.Entry_vendas_TEMA.place(x=25, y=130)
        self.Entry_vendas_descricao = Entry(self.app5, font='times 15', width=15)
        self.Entry_vendas_descricao.place(x=25, y=165)
        self.Entry_vendas_TAMANHO = Entry(self.app5, font='times 15', width=7)
        self.Entry_vendas_TAMANHO.place(x=25, y=200)
        self.Entry_vendas_QUANTIDADE = Entry(self.app5, font='times 15 ', width=7)
        self.Entry_vendas_QUANTIDADE.place(x=25, y=235)
        self.Entry_vendas_VALOR = Entry(self.app5, font='times 15 ', width=7)
        self.Entry_vendas_VALOR.place(x=25, y=270)

        ##################### ENTRYS DA FUNÇAO PESQUISAR SEGUIMENTO ###############################

        self.Entry_vendas_pesquisa = Entry(self.app5, font='times 15 bold', width=20)
        self.Entry_vendas_pesquisa.place(x=395, y=200)

        self.Entry_pesquisa_tema_venda = Entry(self.app5, font='times 15 bold', width=10)
        self.Entry_pesquisa_tema_venda.place(x=395, y=235)

        self.Entry_pesquisa_descricao_venda = Entry(self.app5, font='times 15 bold', width=10)
        self.Entry_pesquisa_descricao_venda.place(x=395, y=270)

        self.lb_vendas_pesquisa_tema = Label(self.app5, text='TEMA', bg='#7f7f7f',
                                             fg='white', relief='raised', font='times 14 bold')
        self.lb_vendas_pesquisa_tema.place(x=510, y=235)

        self.lb_vendas_pesquisa_descricao = Label(self.app5, text='DESCRIÇAO', bg='#7f7f7f',
                                                  fg='white', relief='raised', font='times 14 bold')
        self.lb_vendas_pesquisa_descricao.place(x=510, y=270)

#                                   FRAME E BOTAO DE PESQUISA DA FUNÇAO DELETAR                                     #

        self.lab_vendas = Label(self.app5, text='BUSCA POR CLIENTE', bg='#6b6a68', fg='white', relief='raised',
                                font='times 14')
        self.lab_vendas.place(x=630, y=70)
        self.Entry_venda_cliente = Entry(self.app5, font='times 12')
        self.Entry_venda_cliente.place(x=630, y=105, width=155)

        self.bt_alt_PESQUISAR = Button(self.app5, text='REALIZAR VENDA', bg='#574c26', fg='#ffffff', pady=5,
                                       font='times 9', command=self.realizar_venda)
        self.bt_alt_PESQUISAR.place(x=730, y=265)

        self.data_venda = Label(self.app5, text='DATA DA VENDA', bg='#7f7f7f', fg='white',
                                font='times 15 bold')
        self.data_venda.place(x=320, y=50)

        self.entry_data = Entry(self.app5, font='times 15 bold', width=10)
        self.entry_data.place(x=370, y=100)

        self.bt_vendas = Button(self.app5, text='BUSCAR', bg='#e76e0c', fg='white', font='times 10',
                                command=self.buscar_cliente_sql)
        self.bt_vendas.place(x=790, y=105)

        self.bt_vendas_voltar = Button(self.app5, text='VOLTAR',
                                       command=self.app5.destroy, bg='#7f7f7f', fg='white', relief=FLAT, pady=5,
                                       font='times 15')
        self.bt_vendas_voltar.place(x=730, y=20)

        self.bt_vendas_pesquisar = Button(self.app5, text='PESQUISAR',
                                          command=self.pesquisar_seguimento_venda, bg='#f6f3ea',
                                          fg='black', pady=5,
                                          font='times 10')
        self.bt_vendas_pesquisar.place(x=610, y=195)


        self.total=PhotoImage(file='imagens/total.png')
        self.valor_total= Button(self.app5, text='VALOR A PAGAR',image=self.total,compound=TOP,
                                          command=self.valor_total_venda, bg='#7f7f7f',relief=FLAT,
                                          fg='YELLOW'
                                          )
        self.valor_total.place(x=850, y=160)

        self.img_calendario = PhotoImage(file='imagens/calendario.png')
        self.bt_calendario = Button(self.app5, text='data', command=self.calendario_vendas,font='times 2 bold',
                                    image=self.img_calendario, bg='#7f7f7f', relief=FLAT)
        self.bt_calendario.place(x=460, y=85)

    def calendario_vendas(self):
        self.calendario1 = Calendar(self.app5, fg='gray75', background='darkblue', font='times 9 bold', locale='pt_br')
        self.calendario1.place(x=350, y=90)
        self.data_venda = Button(self.app5, text='INSERIR DATA', bg='#021f53', fg='white',
                                 font='times 10 bold', command=self.inserir_data_venda)
        self.data_venda.place(x=400, y=275)

    def inserir_data_venda(self):
        dataini = self.calendario1.get_date()
        self.calendario1.destroy()
        self.entry_data.delete(0, END)
        self.entry_data.insert(END, dataini)
        self.data_venda.destroy()

    def treeview_vendas_valores(self):
        self.columns = (
            'data', 'cliente', 'seguimento', 'tema', 'descricao', 'tamanho', 'quantidade', 'valor', 'total_venda')
        self.treeview_vendas = ttk.Treeview(self.app5, columns=self.columns, show='headings')
        self.treeview_vendas.heading('data', text='Data')
        self.treeview_vendas.heading('cliente', text='Cliente')
        self.treeview_vendas.heading('seguimento', text='Seguimento')
        self.treeview_vendas.heading('tema', text='Tema')
        self.treeview_vendas.heading('descricao', text='Descriçao')
        self.treeview_vendas.heading('tamanho', text='Tamanho')
        self.treeview_vendas.heading('quantidade', text='Quantidade')
        self.treeview_vendas.heading('valor', text='Valor')
        self.treeview_vendas.heading('total_venda', text='Total')
        self.treeview_vendas.column('cliente', width=80)
        self.treeview_vendas.column('data', width=75)
        self.treeview_vendas.column('seguimento', width=200)
        self.treeview_vendas.column('tema', width=160)
        self.treeview_vendas.column('descricao', width=150)
        self.treeview_vendas.column('tamanho', width=60)
        self.treeview_vendas.column('quantidade', width=70)
        self.treeview_vendas.column('valor', width=50)
        self.treeview_vendas.column('total_venda', width=50)
        self.treeview_vendas.place(x=20, y=330, width=906, height=240)
        self.scroll_vend = Scrollbar(self.app5, orient=VERTICAL, command=self.treeview_vendas.yview)
        self.treeview_vendas.configure(yscrollcommand=self.scroll_vend.set)
        self.scroll_vend.place(x=930, y=330, width=25, height=238)
        self.treeview_vendas.bind("<Double-1>", self.double_click_vendas)

    def buscar_cliente_sql(self):
        self.treeview_vendas.delete(*self.treeview_vendas.get_children())
        self.vendas_Data = self.entry_data.get().strip()
        self.busc_cliente = self.Entry_venda_cliente.get().strip()
        self.sql()

        if self.busc_cliente == '' or self.vendas_Data == '':
            messagebox.showinfo('', 'PREENCHA O CAMPO DATA E INFORME O NOME DO CLIENTE')
        else:
            self.pesquisar_cliente = self.cursor.execute(
            f""" SELECT data_venda,cliente,
                seguimento_venda,
                tema_venda,
                descricao_venda,
                tamanho_venda,
                quantidade_venda,
                valor_venda,
                total_venda FROM seguimento_venda
                WHERE cliente = '{self.busc_cliente}'
				AND data_venda= '{self.vendas_Data}'""")

            for linha in self.cursor.fetchall():
                self.pesqui = linha.data_venda,\
                              linha.cliente,\
                              linha.seguimento_venda,\
                              linha.tema_venda, \
                              linha.descricao_venda,\
                              linha.tamanho_venda,\
                              linha.quantidade_venda,\
                              linha.valor_venda,\
                              linha.total_venda

                self.treeview_vendas.insert('', 'end', values=self.pesqui)

    def limpa_campo_vendas(self):
        self.entry_data.delete(0, END)
        self.Entry_vendas_cliente.delete(0, END)
        self.Entry_vendas_SEGUIMENTO.delete(0, END)
        self.Entry_vendas_TEMA.delete(0, END)
        self.Entry_vendas_descricao.delete(0, END)
        self.Entry_vendas_TAMANHO.delete(0, END)
        self.Entry_vendas_QUANTIDADE.delete(0, END)
        self.Entry_vendas_VALOR.delete(0, END)

    def double_click_vendas(self, event):
        self.treeview_vendas.selection()
        self.limpa_campo_vendas()
        try:
            for i in self.treeview_vendas.selection():
                col1, col2, col3, col4, col5, col6, col7, col8 = self.treeview_vendas.item(i, 'values')
                self.entry_data.insert(END, col1)
                self.Entry_vendas_cliente.insert(END, col2)
                self.Entry_vendas_SEGUIMENTO.insert(END, col3)
                self.Entry_vendas_TEMA.insert(END, col4)
                self.Entry_vendas_descricao.insert(END, col5)
                self.Entry_vendas_TAMANHO.insert(END, col6)
                self.Entry_vendas_QUANTIDADE.insert(END, col7)
                self.Entry_vendas_VALOR.insert(END, col8)
                self.Entry_vendas_QUANTIDADE.delete(0, END)
                self.Entry_vendas_cliente.delete(0, END)
                self.entry_data.delete(0, END)
        except:
            messagebox.showinfo('', 'para alterar os dados do cliente vc devera ir para opc cliente')

    def pesquisar_seguimento_venda(self):
        self.treeview_vendas.delete(*self.treeview_vendas.get_children())
        self.venda_pesquisa = self.Entry_vendas_pesquisa.get().strip()
        self.pesquisa_tema = self.Entry_pesquisa_tema_venda.get().strip()
        self.pesquisa_descricao = self.Entry_pesquisa_descricao_venda.get().strip()

        self.sql()
        if self.venda_pesquisa == '':
            messagebox.showinfo('', 'DIGITE O SEGUIMENTO QUE DESEJA FAZER A VENDA')
        elif self.venda_pesquisa != '' and self.pesquisa_tema != '' and self.pesquisa_descricao != '':
            self.pesquisar_descricao = self.cursor.execute(f"""SELECT data_venda,cliente,seguimento,tema,descricao
                                                             ,tamanho,quantidade,
                                                             valor FROM seguimento_venda 
                                                             WHERE seguimento like '{self.venda_pesquisa}%'
                                                             and tema like'{self.pesquisa_tema}%'
                                                             AND descricao LIKE '{self.pesquisa_descricao}%'
                                                                ORDER BY seguimento,tema,tamanho
                                                                                                       """)
            for linha in self.cursor.fetchall():
                self.pesqui = linha.data_venda, linha.cliente, linha.seguimento, linha.tema, linha.descricao, \
                              linha.tamanho, linha.quantidade, linha.valor
                self.treeview_vendas.insert('', 'end', values=self.pesqui)

        elif self.venda_pesquisa != '' and self.pesquisa_tema != '':
            self.pesquisar_tema = self.cursor.execute(f"""SELECT data_venda,cliente,seguimento,tema,descricao
                                                         ,tamanho,quantidade,
                                                         valor FROM seguimento_venda
                                                         WHERE seguimento like '{self.venda_pesquisa}%'
                                                         and tema like'{self.pesquisa_tema}%'
                                                         ORDER BY seguimento,tema,descricao,tamanho DESC;
                                                                                           """)
            for linha in self.cursor.fetchall():
                self.pesqui = linha.data_venda, linha.cliente, linha.seguimento, linha.tema, linha.descricao, \
                              linha.tamanho, linha.quantidade, linha.valor
                self.treeview_vendas.insert('', 'end', values=self.pesqui)
        else:
            self.pesquisar_seg = self.cursor.execute(f"""SELECT data_venda,cliente,seguimento,tema,descricao
                                                                        ,tamanho,quantidade,
                                                                valor FROM seguimento_venda
                                                                     WHERE seguimento LIKE '{self.venda_pesquisa}%'
                                                                     ORDER BY seguimento,tema,descricao,tamanho DESC;
                                         """)
            for linha in self.cursor.fetchall():
                self.pesqui = linha.data_venda, linha.cliente, linha.seguimento, linha.tema, linha.descricao, \
                              linha.tamanho, linha.quantidade, linha.valor
                self.treeview_vendas.insert('', 'end', values=self.pesqui)
        self.desconectar()

    def realizar_venda(self):
        self.treeview_vendas.delete(*self.treeview_vendas.get_children())
        self.sql()
        self.venda_data = self.entry_data.get()
        self.venda_cliente = self.Entry_vendas_cliente.get()
        self.venda_segui = self.Entry_vendas_SEGUIMENTO.get()
        self.venda_tema = self.Entry_vendas_TEMA.get()
        self.venda_desc = self.Entry_vendas_descricao.get()
        self.venda_tamanho = self.Entry_vendas_TAMANHO.get()
        self.venda_quantidade = self.Entry_vendas_QUANTIDADE.get()
        self.venda_valor = self.Entry_vendas_VALOR.get()
        # try:
        if self.venda_data == '':
            messagebox.showinfo('', 'preencha o campo da data')
        elif self.venda_cliente == '' \
                or self.venda_segui == '' \
                or self.venda_tema == '' \
                or self.venda_desc == '' \
                or self.venda_tamanho == '':
            messagebox.showinfo('', 'PREENCHA OS CAMPOS CORRETAMENTE PARA REAÇIZAR A VENDA!')
        else:
            self.msg = messagebox.askyesno('CLIENTE', f'DESEJA REALIZAR AVENDA PARA {self.venda_cliente}')
            if self.msg == TRUE:
                tabela = self.cursor.execute(f"""SELECT quantidade FROM seguimento_venda WHERE
                                    seguimento = '{self.venda_segui}'AND  tamanho = '{self.venda_tamanho}'
                                                                    AND  tema = '{self.venda_tema}'  """)
                for t in tabela:
                    estoque = t[0]
                    if estoque >= int(self.venda_quantidade):
                        self.registrar = self.cursor.execute(f"""INSERT INTO seguimento_venda
                                                                (data_venda,cliente,seguimento_venda,
                                                                 tema_venda,
                                                                 descricao_venda,
                                                                  tamanho_venda,
                                                                   quantidade_venda,
                                                                    valor_venda )
                                             VALUES
                                                              ('{self.venda_data}',
                                                               '{self.venda_cliente}',
                                                               '{self.venda_segui}',
                                                               '{self.venda_tema}',
                                                               '{self.venda_desc}',
                                                               '{self.venda_tamanho}', 
                                                                {self.venda_quantidade},
                                                                {self.venda_valor}
                                                                                )""")
                        self.vender = self.cursor.execute(f"""UPDATE seguimento_venda SET
                                                     quantidade = quantidade  - {self.venda_quantidade}
                                                      WHERE seguimento =  '{self.venda_segui}'
                                                     AND  tema = '{self.venda_tema}'
                                                     AND tamanho = '{self.venda_tamanho}';
                                                      """)
                        self.pesquisar_venda = self.cursor.execute(
                            f""" SELECT data_venda,
                            cliente,
                            seguimento_venda,
                            tema_venda,
                            descricao_venda,
                            tamanho_venda
                            ,quantidade_venda,
                            valor_venda,
                            total_venda FROM seguimento_venda
                                                       WHERE cliente = '{self.venda_cliente}'
                                                         and data_venda= '{self.venda_data}'
                                                               """)
                        for linha in self.cursor.fetchall():
                            self.pesqui = linha.data_venda, \
                                          linha.cliente, \
                                          linha.seguimento_venda, \
                                          linha.tema_venda, \
                                          linha.descricao_venda, \
                                          linha.tamanho_venda, \
                                          linha.quantidade_venda, \
                                          linha.valor_venda, \
                                          linha.total_venda
                            self.treeview_vendas.insert('', 'end', values=self.pesqui)
                        msg = messagebox.askyesno('', f'DESEJA CONTINUA A VENDA PARA {self.venda_cliente}?')

                        if msg == TRUE:
                            self.Entry_vendas_cliente.delete(0, END)
                            self.Entry_vendas_QUANTIDADE.delete(0, END)
                        else:
                            self.entry_data.delete(0, END)
                            self.Entry_vendas_cliente.delete(0, END)
                            self.Entry_vendas_SEGUIMENTO.delete(0, END)
                            self.Entry_vendas_TEMA.delete(0, END)
                            self.Entry_vendas_descricao.delete(0, END)
                            self.Entry_vendas_TAMANHO.delete(0, END)
                            self.Entry_vendas_QUANTIDADE.delete(0, END)
                            self.Entry_vendas_VALOR.delete(0, END)
                            self.cursor.commit()
                            self.desconectar()

                    else:
                        msg = (f'NAO E POSSIVEL REALIZAR A VENDA, SEU ESTOQUE CONTÉM {estoque} UNIDADE(S)')
                        messagebox.showerror('ESTOQUE INSUFICIENTE', msg)
            self.cursor.commit()
            self.desconectar()

    def valor_total_venda(self):
        self.sql()
        self.venda_cliente = self.Entry_vendas_cliente.get()
        self.venda_data = self.entry_data.get()
        if self.venda_cliente != '' and self.venda_data != '':
            valor = self.cursor.execute(
                f"""SELECT SUM(total_venda) FROM seguimento_venda 
                WHERE cliente = '{self.venda_cliente}' 
                AND data_venda = '{self.venda_data}'""")
            try:
                for p in valor:
                    pagar=p[0]
                    messagebox.showinfo('',f'TOTAL A PAGAR R$: {pagar:,.2f}')
            except:
                messagebox.showwarning('AVISO ! ',f'NENHUM VALOR ENCONTRADO PARA {self.venda_cliente}')
        else:
            messagebox.showinfo('AVISO','INFORME O NOME DO CLIENTE EA DATA PARA CONSULTA O VALOR DE COMPRA')

#                                            FUNÇAO DE ALTERAR VALORES                                                 #

    def alterar_valores(self):
        self.app4 = Toplevel()
        self.app4.transient(self.app)
        self.app4.geometry('940x525+310+110')
        self.app4.resizable(False, False)
        self.app4.focus_force()
        self.app4.grab_set()
        self.app4.configure(bg='#474644')

        self.framealterar = Frame(self.app4, width=920, height=245, borderwidth=2, relief="raised",
                                  background='#7f7f7f', highlightthickness=3, highlightbackground='#04457c')
        self.framealterar.place(x=10,y=10)

        self.frameinferior = Frame(self.app4, bg="white", width=920, height=240, borderwidth=9, relief="raised",
                                   background='#7f7f7f')
        self.frameinferior.place(x=10, y=270)

        self.lb_alterar = Label(self.app4, text='ALTERAR VALORES', bg='#033b6a',
                                fg='#ffffff', font='times 15 bold', relief='raised')
        self.lb_alterar.place(x=220, y=23)

        self.lb_alt_id = Label(self.app4, text='ID', bg='#7f7f7f',
                               fg='white', relief='raised', font='times 15 bold', width=5)
        self.lb_alt_id.place(x=110, y=30)

        self.lb_alt_seguimento = Label(self.app4, text='SEGUIMENTO', bg='#7f7f7f',
                                       fg='white', relief='raised', font='times 15 bold')
        self.lb_alt_seguimento.place(x=190, y=63)

        self.lb_alt_tema = Label(self.app4, text='TEMA ou COR', bg='#7f7f7f',
                                 fg='white', relief='raised', font='times 15 bold')
        self.lb_alt_tema.place(x=210, y=93)

        self.lb_alt_descricao = Label(self.app4, text='DESCRIÇAO', bg='#7f7f7f',
                                      fg='white', relief='raised', font='times 15 bold')
        self.lb_alt_descricao.place(x=190, y=123)

        self.lb_alt_tamanho = Label(self.app4, text='TAMANHO', bg='#7f7f7f',
                                    fg='white', relief='raised', font='times 15 bold')
        self.lb_alt_tamanho.place(x=110, y=153)

        self.lb_alt_quantidade = Label(self.app4, text='QUANTIDADE', bg='#7f7f7f',
                                       fg='white', relief='raised', font='times 15 bold')
        self.lb_alt_quantidade.place(x=110, y=183)

        self.lb_alt_valor = Label(self.app4, text='VALOR', bg='#7f7f7f',
                                  fg='white', relief='raised', font='times 15 bold')
        self.lb_alt_valor.place(x=110, y=215)

        self.treeview_alt_valores()

        # MEUS ENTRYS DA FUNÇAO

        self.Entry_alt_id = Entry(self.app4, font='times 15', width=5)
        self.Entry_alt_id.place(x=30, y=30)
        self.Entry_alt_seguimento = Entry(self.app4, font='times 15', width=15)
        self.Entry_alt_seguimento.place(x=30, y=65)
        self.Entry_alt_tema = Entry(self.app4, font='times 15', width=17)
        self.Entry_alt_tema.place(x=30, y=95)
        self.Entry_alt_descricao = Entry(self.app4, font='times 15', width=15)
        self.Entry_alt_descricao.place(x=30, y=125)
        self.Entry_alt_tamanho = Entry(self.app4, font='times 15', width=7)
        self.Entry_alt_tamanho.place(x=30, y=155)
        self.Entry_alt_quantidade = Entry(self.app4, font='times 15', width=7)
        self.Entry_alt_quantidade.place(x=30, y=185)
        self.Entry_alt_valor = Entry(self.app4, font='times 15 ', width=7)
        self.Entry_alt_valor.place(x=30, y=215)
#                                   LABEL E BOTAO DE PESQUISA DA FUNÇAO DELETAR                                        #

        self.lab_alt = Label(self.app4, text='BUSCA POR SEGUIMENTO', bg='#033b6a', fg='white', relief='raised',
                             font='times 15 bold')
        self.lab_alt.place(x=400, y=70)

        self.lab_tema = Label(self.app4, text='TEMA', bg='#033b6a', fg='white', relief='raised',
                              font='times 15 bold')
        self.lab_tema.place(x=575, y=130)

        #                                 ENTRYS                                                                       #
        self.Entry_alt_pesquisar = Entry(self.app4, font='times 15 bold', width=7)
        self.Entry_alt_pesquisar.place(x=420, y=100, width=210)
        self.Entry_alt_temaa = Entry(self.app4, font='times 15 bold', width=7)
        self.Entry_alt_temaa.place(x=420, y=130, width=150)

        #                                     BOTOES

        self.bt_alt_pesquisar = Button(self.app4, text='PESQUISAR', bg='#1481dd', fg='white',
                                       font='times 15 bold', command=self.pesquisar_seguimento_alt)
        self.bt_alt_pesquisar.place(x=560, y=200)

        self.bt_alt_deletar = Button(self.app4, text='ALTERAR', bg='#04457c', fg='white',
                                     font='times 15 bold', command=self.alterar_os_valores_sql)
        self.bt_alt_deletar.place(x=400, y=200)

        self.bt_alt_voltar = Button(self.app4, text='VOLTAR', font='times 18 bold', relief=FLAT, bg='#7f7f7f',
                                    fg='white',
                                    command=self.app4.destroy)
        self.bt_alt_voltar.place(x=580, y=20)

        #                               treeview da funçao deletar                                                     #

    def treeview_alt_valores(self):
        self.columns = ('id', 'seguimento', 'tema', 'descricao', 'tamanho', 'quantidade', 'valor')
        self.treeview_alt = ttk.Treeview(self.app4, columns=self.columns, show='headings')
        self.treeview_alt.heading('id', text='ID')
        self.treeview_alt.heading('seguimento', text='Seguimento')
        self.treeview_alt.heading('tema', text='Tema')
        self.treeview_alt.heading('descricao', text='Descriçao')
        self.treeview_alt.heading('tamanho', text='Tamanho')
        self.treeview_alt.heading('quantidade', text='Quantidade')
        self.treeview_alt.heading('valor', text='Valor')

        self.treeview_alt.column('id', width=20)

        self.treeview_alt.column('seguimento', width=200)
        self.treeview_alt.column('tema', width=120)
        self.treeview_alt.column('descricao', width=140)
        self.treeview_alt.column('tamanho', width=10)
        self.treeview_alt.column('quantidade', width=25)
        self.treeview_alt.column('valor', width=20)
        self.scrollalt = Scrollbar(self.app4, orient=VERTICAL, command=self.treeview_alt.yview)
        self.treeview_alt.configure(yscrollcommand=self.scrollalt.set)
        self.scrollalt.place(x=893, y=290, width=25, height=208)
        self.treeview_alt.place(x=20, y=290, width=870, height=210)
        self.treeview_alt.bind("<Double-1>", self.double_click_alt)

    def inserir_daddos_alt(self):
        self.alt_id = self.Entry_alt_id.get().strip()
        self.alt_segui = self.Entry_alt_seguimento.get().strip()
        self.alt_tema = self.Entry_alt_tema.get().strip()
        self.alt_desc = self.Entry_alt_descricao.get().strip()
        self.alt_tamanho = self.Entry_alt_tamanho.get().strip()
        self.alt_quantidade = self.Entry_alt_quantidade.get().strip()
        self.alt_valor = self.Entry_alt_valor.get().strip()
        self.alt_VALORES = self.alt_id, self.alt_segui, self.alt_tema, self.alt_desc, self.alt_tamanho,\
                           self.alt_quantidade, self.alt_valor
        self.treeview_alt.insert('', 'end', values=self.alt_VALORES)

    def limpa_campo_alt(self):
        self.Entry_alt_id.delete(0, END)
        self.Entry_alt_seguimento.delete(0, END)
        self.Entry_alt_tema.delete(0, END)
        self.Entry_alt_descricao.delete(0, END)
        self.Entry_alt_tamanho.delete(0, END)
        self.Entry_alt_quantidade.delete(0, END)
        self.Entry_alt_valor.delete(0, END)

    def double_click_alt(self, event):
        self.treeview_alt.selection()
        self.limpa_campo_alt()
        for i in self.treeview_alt.selection():
            col1, col2, col3, col4, col5, col6, col7 = self.treeview_alt.item(i, 'values')
            self.Entry_alt_id.insert(END, col1)
            self.Entry_alt_seguimento.insert(END, col2)
            self.Entry_alt_tema.insert(END, col3)
            self.Entry_alt_descricao.insert(END, col4)
            self.Entry_alt_tamanho.insert(END, col5)
            self.Entry_alt_quantidade.insert(END, col6)
            self.Entry_alt_valor.insert(END, col7)
            self.Entry_alt_valor.delete(0, END)

    def pesquisar_seguimento_alt(self):
        self.sql()
        self.treeview_alt.delete(*self.treeview_alt.get_children())
        self.alt_id = self.Entry_alt_id.get()
        self.alt_pesquisar = self.Entry_alt_pesquisar.get()
        self.alt_pesquisar_tema = self.Entry_alt_temaa.get()

        if self.alt_pesquisar == '':
            messagebox.showinfo('', 'DIGITE O SEGUIMENTO QUE DESEJA ALTERA O VALOR')
        elif self.alt_pesquisar != '' and self.alt_pesquisar_tema != '':
            self.pesquisar_Tema = self.cursor.execute(f"""SELECT id,seguimento,
                                                                    tema,
                                                                    descricao,
                                                                    tamanho,
                                                                    quantidade,
                                                                    valor FROM seguimento_venda
                                                                     WHERE seguimento LIKE '{self.alt_pesquisar}%'
                                                                     AND tema LIKE '{self.alt_pesquisar_tema}%'
                                                                                                  """)

            for linha in self.cursor.fetchall():
                self.pesquisa_alt = linha.id, \
                                    linha.seguimento, \
                                    linha.tema, \
                                    linha.descricao, \
                                    linha.tamanho, \
                                    linha.quantidade, \
                                    linha.valor
                self.treeview_alt.insert('', 'end', values=self.pesquisa_alt)
            self.desconectar()

        else:
            self.pesquisar = self.cursor.execute(f"""SELECT id,seguimento,
                                                          tema,descricao,tamanho,quantidade,valor FROM seguimento_venda
                                                          WHERE seguimento LIKE '{self.alt_pesquisar}%'
    
                                                                                       """)

            for linha in self.cursor.fetchall():
                self.pesquisa_alt = linha.id, \
                                    linha.seguimento, \
                                    linha.tema, \
                                    linha.descricao, \
                                    linha.tamanho, \
                                    linha.quantidade, \
                                    linha.valor
                self.treeview_alt.insert('', 'end', values=self.pesquisa_alt)
            self.desconectar()

    def alterar_os_valores_sql(self):
        self.sql()
        self.alt_segui = self.Entry_alt_seguimento.get().strip()
        self.alt_tema = self.Entry_alt_tema.get().strip()
        self.alt_quanti=self.Entry_alt_quantidade.get().strip()
        self.alt_desc = self.Entry_alt_descricao.get().strip()
        self.alt_tamanho = self.Entry_alt_tamanho.get().strip()
        self.alt_valor = self.Entry_alt_valor.get().strip().replace(',', '.')
        try:
            if self.alt_segui == \
                    '' or self.alt_tema == '' or \
                    self.alt_desc == '' or \
                    self.alt_tamanho == '' or self.alt_quanti == ''or self.alt_valor == '':
                messagebox.showinfo('', 'PREENCHA OS CAMPOS PARA REALIZAR A ALTERAÇAO NO VALOR')

            else:
                self.Msg = messagebox.askyesno(
                    '', f'DESEJA RALIZAR A ALTERAÇAO NO VALOR DO SEGUIMENTO  {self.alt_segui}')
                if self.Msg == TRUE:
                    self.Msg1 = messagebox.askyesno('', f'O VALOR DO SEGUIMENTO   {self.alt_segui}    \n'
                                                        f'SERA ALTERADO PARA {self.alt_valor} R$ DESEJA CONTINUAR?')
                    if self.Msg1 == TRUE:
                        self.alterar_valorSQL = self.cursor.execute(
                            f"""UPDATE seguimento_venda SET valor = '{self.alt_valor}'
                                                        WHERE seguimento = '{self.alt_segui}'
                                                        AND tema = '{self.alt_tema}'
                                                        AND descricao = '{self.alt_desc}'
                                                        AND tamanho = '{self.alt_tamanho}'
                                                        AND quantidade= '{self.alt_quanti}'""")

                        messagebox.showinfo('', 'O VALOR FOI ALTERADO COM SUCESSO')
                        self.cursor.commit()
                        self.desconectar()
                    else:
                        messagebox.showinfo('', 'NENHUM VALOR SERA ALTERADO')
                else:
                    messagebox.showinfo('', 'NENHUM VALOR SERA ALTERADO')
        except:
            messagebox.showerror(
                'ERRO !',f'NÃO E POSSIVEL ALTERA O VALOR {self.alt_valor} VERIFIQUE SE INFORMOU CORRETAMENTE O VALOR')

#                                               FUNÇAO DELETAR ITENS                                      #
    # label e entrys da funçao deletar tabela
    def deletar_item_tabela(self):
        self.app3 = Toplevel()
        self.app3.geometry('910x560+300+90')
        self.app3.resizable(False, False)
        self.app3.transient(self.app)
        self.app3.focus_force()
        self.app3.grab_set()

        self.framePesquisa = Frame(self.app3, bg="white", width=890, height=270, borderwidth=9, relief="raised",
                                   background='#7f7f7f')
        self.framePesquisa.place(x=10, y=10)
        self.frameDeletar = Frame(self.app3, bg="white", width=890, height=260, borderwidth=9, relief="raised",
                                  background='#7f7f7f')
        self.frameDeletar.place(x=10, y=290)
        self.lbDeletar = Label(self.app3, text='ÁREA DE EXCLUSAO DE DADOS', bg='red',
                               fg='white', font='times 15', relief='raised')
        self.lbDeletar.place(x=320, y=23)
        self.lb_id = Label(self.app3, text='ID', bg='#7f7f7f',
                           fg='white', relief='raised', font='times 15', width=5)
        self.lb_id.place(x=110, y=30)
        self.lb_seguimento = Label(self.app3, text='SEGUIMENTO', bg='#7f7f7f',
                                fg='white', relief='raised', font='times 15', width=15)
        self.lb_seguimento.place(x=180, y=63)

        self.lb_tema = Label(self.app3, text='TEMA', bg='#7f7f7f',
                                   fg='white', relief='raised', font='times 15', width=15)
        self.lb_tema.place(x=180, y=97)

        self.lb_descricao = Label(self.app3, text='DESCRIÇAO', bg='#7f7f7f',
                             fg='white', relief='raised', font='times 15', width=15)
        self.lb_descricao.place(x=180, y=132)

        self.lb_tamanho = Label(self.app3, text='TAMANHO', bg='#7f7f7f',
                                fg='white', relief='raised', font='times 15', width=15)
        self.lb_tamanho.place(x=105, y=170)

        self.lb_quantidade = Label(self.app3, text='QUANTIDADE', bg='#7f7f7f',
                                   fg='white', relief='raised', font='times 15', width=15)
        self.lb_quantidade.place(x=105, y=203)

        self.lb_valor = Label(self.app3, text='VALOR', bg='#7f7f7f',
                              fg='white', relief='raised', font='times 15', width=15)
        self.lb_valor.place(x=105, y=240)
        self.treeview_deletar_seguimento()

# MEUS ENTRYS DA FUNÇAO DELETAR
        self.entry_del_id = Entry(self.app3, font='times 15', width=7)
        self.entry_del_id.place(x=20, y=30)
        self.entry_del_seguimento = Entry(self.app3, font='times 15', width=15)
        self.entry_del_seguimento.place(x=20, y=63)
        self.entry_del_tema = Entry(self.app3, font='times 15', width=15)
        self.entry_del_tema.place(x=20, y=98)
        self.entry_del_descricao = Entry(self.app3, font='times 15', width=15)
        self.entry_del_descricao.place(x=20, y=133)
        self.entry_del_tamanho = Entry(self.app3, font='times 15', width=7)
        self.entry_del_tamanho.place(x=20, y=168)
        self.entry_del_quantidade = Entry(self.app3, font='times 15', width=7)
        self.entry_del_quantidade.place(x=20, y=203)
        self.entry_del_valor = Entry(self.app3, font='times 15', width=7)
        self.entry_del_valor.place(x=20, y=240)

#                                           FRAME E BOTAO DE PESQUISA DA FUNÇAO DELETAR                               #
        self.lab_del_pesquisa = Label(self.app3, text='BUSCA POR SEGUIMENTO', bg='#7f7f7f',
                                                                              fg='white', relief='raised',
                                                                              font='times 15')
        self.lab_del_pesquisa.place(x=560, y=70)

        self.lab_del_TEMA = Label(self.app3, text='TEMA', bg='#7f7f7f',
                                                          fg='white', relief='raised',
                                                          font='times 15')
        self.lab_del_TEMA.place(x=670, y=130)

        #                                         ENTRYS   PESQUISA   DELETA                                         #
        self.entry_del_pesquisa = Entry(self.app3, font='times 15', width=15)
        self.entry_del_pesquisa.place(x=560, y=100, width=160)
        self.entry_del_temaa = Entry(self.app3, font='times 15', width=15)
        self.entry_del_temaa.place(x=560, y=130, width=100)
#                                                    BOTOES
        self.bt_del_pesquisa = Button(self.app3, text='PESQUISAR', bg='#0f92f8', fg='white',
                                      font='times 11', command=self.del_pesquisar_sql)
        self.bt_del_pesquisa.place(x=735, y=100)

        self.bt_del_deletar = Button(self.app3, text='DELETAR', bg='#af0000', fg='white', font='times 17',
                                     command=self.deletar_item_sql)
        self.bt_del_deletar.place(x=395, y=100)

        self.bt_del_pesquisa = Button(self.app3, text='CANCELAR', bg='#032812', fg='white', font='times 11',
                                      height=2, command=self.app3.destroy)
        self.bt_del_pesquisa.place(x=600, y=220)

    def inserir_daddos_deletar(self):
        self.del_id = self.entry_del_id.get().strip()
        self.del_seguimento = self.entry_del_seguimento.get().strip()
        self.del_tema = self.entry_del_tema.get().strip()
        self.del_descricao = self.entry_del_descricao.get().strip()
        self.del_tamanho = self.entry_del_tamanho.get().strip()
        self.del_quantidade = self.entry_del_quantidade.get().strip()
        self.del_valor = self.entry_del_valor.get()
        #                               treeview da funçao deletar                                                     #

    def treeview_deletar_seguimento(self):
        self.columns = ('id', 'seguimento', 'tema', 'descricao', 'tamanho', 'quantidade', 'valor')
        self.treeview_deletar = ttk.Treeview(self.app3, columns=self.columns, show='headings')
        self.treeview_deletar.heading('id', text='ID')
        self.treeview_deletar.heading('seguimento', text='seguimento')
        self.treeview_deletar.heading('tema', text='Tema')
        self.treeview_deletar.heading('descricao', text='Descriçao')
        self.treeview_deletar.heading('tamanho', text='Tamanho')
        self.treeview_deletar.heading('quantidade', text='Quantidade')
        self.treeview_deletar.heading('valor', text='Valor')

        self.treeview_deletar.column('id', width=40)
        self.treeview_deletar.column('seguimento', width=80)
        self.treeview_deletar.column('tema', width=100)
        self.treeview_deletar.column('descricao', width=180)
        self.treeview_deletar.column('tamanho', width=50)
        self.treeview_deletar.column('quantidade', width=50)
        self.treeview_deletar.column('valor', width=50)
        self.treeview_deletar.place(x=20, y=300, width=840, height=240)
        self.scrolldel = Scrollbar(self.app3, orient=VERTICAL, command=self.treeview_deletar.yview)
        self.treeview_deletar.configure(yscrollcommand=self.scrolldel.set)
        self.scrolldel.place(x=865, y=300, width=25, height=238)
        self.treeview_deletar.bind("<Double-1>", self.double_click_DELET)

    def limpacampoDELETAR(self):
        self.entry_del_id.delete(0, END)
        self.entry_del_seguimento.delete(0, END)
        self.entry_del_tema.delete(0, END)
        self.entry_del_descricao.delete(0, END)
        self.entry_del_tamanho.delete(0, END)
        self.entry_del_quantidade.delete(0, END)
        self.entry_del_valor.delete(0, END)

    def double_click_DELET(self, event):
        self.treeview_deletar.selection()
        self.limpacampoDELETAR()
        for i in self.treeview_deletar.selection():
            col1, col2, col3, col4, col5, col6, col7 = self.treeview_deletar.item(i, 'values')
            self.entry_del_id.insert(END, col1)
            self.entry_del_seguimento.insert(END, col2)
            # self.entry_del_tema.insert(END, col3)
            self.entry_del_descricao.insert(END, col4)
            self.entry_del_tamanho.insert(END, col5)
            self.entry_del_quantidade.insert(END, col6)
            self.entry_del_valor.insert(END, col7)

    def inserir_dados_doubleClick_delet(self):
        self.treeview_deletar.insert('', END)
        self.inserir_daddos_deletar()

    def del_pesquisar_sql(self):
        self.treeview_deletar.delete(*self.treeview_deletar.get_children())
        self.sql()
        self.bus = self.entry_del_pesquisa.get().strip()
        self.busc_tema = self.entry_del_temaa.get().strip()
        if self.bus == '':
            messagebox.showinfo('', 'E NESSESSARIO INFORMA O SEGUIMENTO QUE DESEJA REALIZAR A BUSCA ')
        elif self.bus != '' and self.busc_tema != '':
            self.pesqui_tema = self.cursor.execute(
                f"""SELECT id,seguimento,tema,descricao,tamanho,quantidade,valor 
                           FROM seguimento_venda
                           WHERE seguimento LIKE '{self.bus}%'
                           AND tema LIKE '{self.busc_tema}%'
                                                """)
            for linha in self.cursor.fetchall():
                self.buscar2 = linha.id, linha.seguimento, linha.tema, linha.descricao, linha.tamanho, \
                               linha.quantidade, linha.valor
                self.treeview_deletar.insert('', 'end', values=self.buscar2)

        elif self.bus != '':
            self.pesqui_tema = self.cursor.execute(
                f"""SELECT id,seguimento,tema,descricao,tamanho,quantidade,valor 
                           FROM seguimento_venda
                           WHERE seguimento LIKE '{self.bus}%'
                                                """)
            for linha in self.cursor.fetchall():
                self.buscar2 = linha.id, linha.seguimento, linha.tema, linha.descricao, linha.tamanho, \
                               linha.quantidade, linha.valor
                self.treeview_deletar.insert('', 'end', values=self.buscar2)

        else:
            self.pesquisar = self.cursor.execute(
                f"""SELECT id,seguimento,tema,descricao,tamanho,quantidade,valor 
                FROM seguimento_venda
                WHERE seguimento = '{self.bus}';
                                  """)
        for linha in self.cursor.fetchall():
            self.buscar2 = linha.id, linha.seguimento, linha.tema, linha.descricao, linha.tamanho,\
                           linha.quantidade, linha.valor
            self.treeview_deletar.insert('', 'end', values=self.buscar2)

    def deletar_item_sql(self):
        self.sql()
        self.dele_id = self.entry_del_id.get().strip()
        self.dele_seguimento = self.entry_del_seguimento.get().strip()
        self.dele_tema = self.entry_del_tema.get().strip()
        self.dele_descricao = self.entry_del_descricao.get().strip()
        self.dele_tamanho = self.entry_del_tamanho.get().strip()
        self.dele_quantidade = self.entry_del_quantidade.get().strip()
        self.dele_valor = self.entry_del_valor.get().strip()
        try:
            if self.dele_seguimento == '' \
                    or self.dele_tema == '' \
                    or self.dele_descricao == '' \
                    or self.dele_tamanho == '':
                messagebox.showinfo('', 'PRENCHA CORRETAMENTE OS CAMPOS QUE DESEJA EXCLUIR')
            else:
                self.MSG = messagebox.askyesno('', f'DESEJA DELETAR O SEGUIMENTO {self.dele_seguimento}?')
            if self.MSG == TRUE:
                self.MSG1 = messagebox.askyesno('', f'O SEGUIMENTO {self.dele_seguimento} SERÁ DELETADO PERMANENTIMENTE'
                                                    ' \n DESEJA CONTINUAR?')
                if self.MSG1 == TRUE:
                    self.delet = self.cursor.execute(f"""DELETE FROM seguimento_venda
                                                          WHERE seguimento = '{self.dele_seguimento}' 
                                                        AND descricao = '{self.dele_descricao}'
                                                        AND id = '{self.dele_id}'
                                                        AND tema = '{self.dele_tema}'
                                                        AND tamanho  = '{self.dele_tamanho}'
                                                                                    """)

                    messagebox.showinfo('', f'O SEGUIMENTO {self.dele_seguimento} FOI DELETADO COM SUCESSO')

                    self.cursor.commit()
                    self.desconectar()
                else:
                    messagebox.showinfo('', 'NENHUM DOS DADOS OBTIDOS FOI DELETADO')
            else:
                messagebox.showinfo('', 'NENHUM DOS DADOS OBTIDOS FOI DELETADO')
        except:
            messagebox.showerror('AVISO','PARA EXCLUIR O SEGUIMENTO E NECESSARIO INFORMAR\n'
                                           'O TEMA DESEJADO CORRESPONDENTE AO SEGUIMENTO QUE DESEJA EXCLUIR')

    #                           CONSULTA DE VENDAS                                                                     #

    def consultar_vendas(self):
        self.app6 = Toplevel()
        self.app6.geometry('970x600+300+56')
        self.app6.resizable(False, False)
        self.app6.transient(self.app)
        self.app6.focus_force()
        self.app6.grab_set()
        self.app6.configure(bg='#2a2a29')

        self.frameconsulta = Frame(self.app6, bg="white", width=955, height=300, borderwidth=9, relief="raised",
                                   background='#7f7f7f')
        self.frameconsulta.place(x=10, y=10)

        self.frameinferiorconculta = Frame(self.app6, bg="white", width=955, height=270, borderwidth=9, relief="raised",
                                           background='#7f7f7f')
        self.frameinferiorconculta.place(x=10, y=315)

        self.lb_consulta = Label(self.app6, text='REALIZAR CONSULTA DE VENDAS ', bg='#574c26',
                                 fg='white', font='times 15', relief='raised')
        self.lb_consulta.place(x=300, y=23)

        self.lb_consulta_cliente = Label(self.app6, text='CLIENTE', bg='#7f7f7f',
                                         fg='white', relief='raised', font='times 14')
        self.lb_consulta_cliente.place(x=190, y=58)

        self.lb_consulta_seguimento = Label(self.app6, text='SEGUIMENTO', bg='#7f7f7f',
                                            fg='white', relief='raised', font='times 14')
        self.lb_consulta_seguimento.place(x=190, y=93)

        self.lb_consulta_tema = Label(self.app6, text='TEMA ou COR', bg='#7f7f7f',
                                      fg='white', relief='raised', font='times 14')
        self.lb_consulta_tema.place(x=190, y=123)

        self.lb_consulta_descricao = Label(self.app6, text='DESCRIÇAO', bg='#7f7f7f',
                                           fg='white', relief='raised', font='times 14')
        self.lb_consulta_descricao.place(x=190, y=165)

        self.lb_consulta_tamanho = Label(self.app6, text='TAMANHO', bg='#7f7f7f',
                                         fg='white', relief='raised', font='times 14')
        self.lb_consulta_tamanho.place(x=110, y=200)

        self.lb_consulta_quantidade = Label(self.app6, text='QUANTIDADE', bg='#7f7f7f',
                                            fg='white', relief='raised', font='times 14')
        self.lb_consulta_quantidade.place(x=110, y=235)

        self.lb_consulta_valor = Label(self.app6, text='VALOR', bg='#7f7f7f',
                                       fg='white', relief='raised', font='times 14')
        self.lb_consulta_valor.place(x=110, y=270)

        self.lb_consulta_total = Label(self.app6, text='TOTAL', bg='#7f7f7f',
                                       fg='white', relief=FLAT, font='times 15 bold')
        self.lb_consulta_total.place(x=270, y=270)
        self.treview_consulta()

        self.Entry_consultar_cliente = Entry(self.app6, font='times 15', width=15)
        self.Entry_consultar_cliente.place(x=25, y=58)
        self.Entry_consultar_seguimento = Entry(self.app6, font='times 15', width=15)
        self.Entry_consultar_seguimento.place(x=25, y=95)
        self.Entry_consultar_tema = Entry(self.app6, font='times 15', width=15)
        self.Entry_consultar_tema.place(x=25, y=130)
        self.Entry_consultar_descricao = Entry(self.app6, font='times 15', width=15)
        self.Entry_consultar_descricao.place(x=25, y=165)
        self.Entry_consultar_tamanho = Entry(self.app6, font='times 15', width=7)
        self.Entry_consultar_tamanho.place(x=25, y=200)
        self.Entry_consultar_quantidade = Entry(self.app6, font='times 15 ', width=7)
        self.Entry_consultar_quantidade.place(x=25, y=235)
        self.Entry_consultar_valor = Entry(self.app6, font='times 15 ', width=7)
        self.Entry_consultar_valor.place(x=25, y=270)
        self.Entry_consultar_total = Entry(self.app6, font='times 15 bold', width=7,
                                           bg='#7f7f7f', fg='white', relief=FLAT)
        self.Entry_consultar_total.place(x=190, y=270)

        #                     LABEL BUSCA POR CLIENTE  DA TELA DE DELETAR                                           #

        self.lab_cliente = Label(self.app6, text='BUSCA POR CLIENTE', bg='#6b6a68', fg='white', relief='raised',
                                 font='times 14')
        self.lab_cliente.place(x=630, y=70)

        self.lab_seguimento = Label(self.app6, text='SEGUIMENTO', bg='#6b6a68', fg='white', relief='raised',
                                    font='times 14')
        self.lab_seguimento.place(x=630, y=130)

        self.lab_tema = Label(self.app6, text='TEMA', bg='#6b6a68', fg='white', relief='raised',
                              font='times 14')
        self.lab_tema.place(x=735, y=185)

        self.Entry_consultar_clientee = Entry(self.app6, font='times 12')
        self.Entry_consultar_clientee.place(x=630, y=105, width=155)

        self.Entry_consultar_seguimentoo = Entry(self.app6, font='times 12')
        self.Entry_consultar_seguimentoo.place(x=630, y=155, width=155)

        self.Entry_consultar_temaa = Entry(self.app6, font='times 12')
        self.Entry_consultar_temaa.place(x=630, y=190, width=100)
        #                               CALENDARIO                                                                     #

        self.data_consultar = Label(self.app6, text='DATA DA VENDA', bg='#7f7f7f', fg='white',
                                    font='times 15 bold')
        self.data_consultar.place(x=385, y=80)

        self.entry_dataa = Entry(self.app6, font='times 15 bold', width=10)
        self.entry_dataa.place(x=390, y=110)


        self.img_calendario = PhotoImage(file='imagens/calendario.png')
        self.bt_consultar_venda = Button(self.app6, text='data', command=self.calendario_consulta,
                                         image=self.img_calendario, bg='#7f7f7f', relief=FLAT)
        self.bt_consultar_venda.place(x=490, y=100)
#                                            BOTOES DA FUNÇAO CONSULTA VENDAS                                          #
        self.bt_consultar = Button(self.app6, text='BUSCAR', bg='#e76e0c', fg='white', font='times 10',
                                            command=self.consulta_cliente_sql)
        self.bt_consultar.place(x=790, y=105)

        self.bt_consultar_voltar = Button(self.app6, text='VOLTAR', command=self.app6.destroy, bg='#7f7f7f', fg='white',
                                          relief=FLAT, pady=5,
                                          font='times 15')
        self.bt_consultar_voltar.place(x=730, y=20)

        self.bt_consultar_excluir = Button(self.app6, text='EXCLUIR VENDA', bg='red', fg='white', pady=2,
                                           font='times 10 bold', command=self.excluir_venda)
        self.bt_consultar_excluir.place(x=410, y=265)

        self.bt_consultar_PESQUISAR = Button(self.app6, text='CORRIGIR VALOR', bg='#f6f3ea', pady=2,
                                             font='times 10', command=self.corrigir_valor_consulta)
        self.bt_consultar_PESQUISAR.place(x=700, y=270)

        self.bt_consultar_adiciona = Button(self.app6, text='ADICIONAR', bg='#f6f3ea', pady=2,
                                            font='times 10', command=self.adicionar_iten)
        self.bt_consultar_adiciona.place(x=550, y=270)

    #                                 FUNÇOES DO CALENDARIO                                                           #

    def calendario_consulta(self):
        self.calendario2 = Calendar(self.app6, fg='gray75', bg='blue', font='times 9 bold', locale='pt_br')
        self.calendario2.place(x=350, y=90)
        self.data_consultar = Button(self.app6, text='INSERIR DATA', bg='#021f53', fg='white',
                                     font='times 10 bold', command=self.inserir_data_consulta)
        self.data_consultar.place(x=400, y=275)

    def inserir_data_consulta(self):
        dataini = self.calendario2.get_date()
        self.calendario2.destroy()
        self.entry_dataa.delete(0, END)
        self.entry_dataa.insert(END, dataini)
        self.data_consultar.destroy()

#                                 FUNÇAO DO BOTAO CONSULTA VENDA                                                    #

    def treview_consulta(self):
        self.columns = (
            'data', 'cliente', 'seguimento', 'tema', 'descricao', 'tamanho', 'quantidade', 'valor', 'total_venda')
        self.treeview_consulta = ttk.Treeview(self.app6, columns=self.columns, show='headings')
        self.treeview_consulta.heading('data', text='Data')
        self.treeview_consulta.heading('cliente', text='Cliente')
        self.treeview_consulta.heading('seguimento', text='Seguimento')
        self.treeview_consulta.heading('tema', text='Tema')
        self.treeview_consulta.heading('descricao', text='Descriçao')
        self.treeview_consulta.heading('tamanho', text='Tamanho')
        self.treeview_consulta.heading('quantidade', text='Quantidade')
        self.treeview_consulta.heading('valor', text='Valor')
        self.treeview_consulta.heading('total_venda', text='Total')

        self.treeview_consulta.column('cliente', width=80)
        self.treeview_consulta.column('data', width=75)
        self.treeview_consulta.column('seguimento', width=200)
        self.treeview_consulta.column('tema', width=130)
        self.treeview_consulta.column('descricao', width=130)
        self.treeview_consulta.column('tamanho', width=60)
        self.treeview_consulta.column('quantidade', width=70)
        self.treeview_consulta.column('valor', width=60)
        self.treeview_consulta.column('total_venda', width=60)
        self.treeview_consulta.place(x=20, y=330, width=905, height=240)
        self.scrollcons = Scrollbar(self.app6, orient=VERTICAL, command=self.treeview_consulta.yview)
        self.treeview_consulta.configure(yscrollcommand=self.scrollcons.set)
        self.scrollcons.place(x=930, y=330, width=25, height=238)
        self.treeview_consulta.bind("<Double-1>", self.double_click_consulta)

    def consulta_cliente_sql(self):
        self.treeview_consulta.delete(*self.treeview_consulta.get_children())
        self.consulta_data = self.entry_dataa.get().strip()
        self.consultar_cliente = self.Entry_consultar_clientee.get().strip()
        self.consulta_seguimentoo = self.Entry_consultar_seguimentoo.get().strip()
        self.consulta_temaa = self.Entry_consultar_temaa.get().strip()

        self.sql()
        try:
            if self.consultar_cliente == '':
                messagebox.showinfo('', 'INFORME O NOME DO CLIENTE')
            elif self.consultar_cliente != '' and \
                    self.consulta_data != '' and \
                    self.consulta_seguimentoo != '' and \
                    self.consulta_temaa != '':
                self.buscaPESQUISA2 = self.cursor.execute(f"""
                    select data_venda,
                    cliente,
                    seguimento_venda,
                    tema_venda,
                    descricao_venda,
                    tamanho_venda,
                    quantidade_venda,
                    valor_venda,
                    total_venda 
                                    FROM seguimento_venda where cliente= '{self.consultar_cliente}'
                                    AND data_venda = '{self.consulta_data}'
                                    AND seguimento_venda LIKE '{self.consulta_seguimentoo}%'
                                    AND tema_venda LIKE '{self.consulta_temaa}%'""")
                for linha in self.cursor.fetchall():
                    self.pesqui = linha.data_venda, \
                                  linha.cliente, \
                                  linha.seguimento_venda, \
                                  linha.tema_venda, \
                                  linha.descricao_venda, \
                                  linha.tamanho_venda, \
                                  linha.quantidade_venda, \
                                  linha.valor_venda, \
                                  linha.total_venda
                    self.treeview_consulta.insert('', 'end', values=self.pesqui)

            elif self.consultar_cliente != '' and self.consulta_data != '' and self.consulta_seguimentoo != '':

                self.buscaPESQUISAa = self.cursor.execute(f"""
                SELECT data_venda,
                cliente,
                seguimento_venda,
                tema_venda,
                descricao_venda,
                tamanho_venda,
                quantidade_venda,
                valor_venda,
                total_venda 
                                FROM seguimento_venda where cliente= '{self.consultar_cliente}'
                                AND data_venda = '{self.consulta_data}'
                                AND seguimento_venda LIKE'{self.consulta_seguimentoo}%'
                                           """)
                for linha in self.cursor.fetchall():
                    self.pesqui = linha.data_venda, \
                                  linha.cliente, \
                                  linha.seguimento_venda, \
                                  linha.tema_venda, \
                                  linha.descricao_venda, \
                                  linha.tamanho_venda, \
                                  linha.quantidade_venda, \
                                  linha.valor_venda, \
                                  linha.total_venda
                    self.treeview_consulta.insert('', 'end', values=self.pesqui)
            else:
                self.buscaPESQUISA = self.cursor.execute(f"""
                select data_venda,
                cliente,
                seguimento_venda,
                tema_venda,
                descricao_venda,
                tamanho_venda,quantidade_venda,valor_venda,total_venda 
                FROM seguimento_venda where cliente= '{self.consultar_cliente}'
                                    and data_venda = '{self.consulta_data}'""")
                for linha in self.cursor.fetchall():
                    self.pesqui = linha.data_venda, \
                                  linha.cliente, \
                                  linha.seguimento_venda, \
                                  linha.tema_venda, \
                                  linha.descricao_venda, \
                                  linha.tamanho_venda, \
                                  linha.quantidade_venda, \
                                  linha.valor_venda, \
                                  linha.total_venda
                    self.treeview_consulta.insert('', 'end', values=self.pesqui)
        except:
            messagebox.showerror('ERRO!!!', 'INFORME OS DADOS CORRETAMENTE')

    def limpa_campo_consulta(self):
        self.entry_dataa.delete(0, END)
        self.Entry_consultar_cliente.delete(0, END)
        self.Entry_consultar_seguimento.delete(0, END)
        self.Entry_consultar_tema.delete(0, END)
        self.Entry_consultar_descricao.delete(0, END)
        self.Entry_consultar_tamanho.delete(0, END)
        self.Entry_consultar_quantidade.delete(0, END)
        self.Entry_consultar_valor.delete(0, END)
        self.Entry_consultar_total.delete(0, END)

    def double_click_consulta(self, event):
        self.treeview_consulta.selection()
        self.limpa_campo_consulta()
        for i in self.treeview_consulta.selection():
            col1, col2, col3, col4, col5, col6, col7, col8, col9 = self.treeview_consulta.item(i, 'values')
            self.entry_dataa.insert(END, col1)
            self.Entry_consultar_cliente.insert(END, col2)
            self.Entry_consultar_seguimento.insert(END, col3)
            self.Entry_consultar_tema.insert(END, col4)
            self.Entry_consultar_descricao.insert(END, col5)
            self.Entry_consultar_tamanho.insert(END, col6)
            self.Entry_consultar_quantidade.insert(END, col7)
            self.Entry_consultar_valor.insert(END, col8)
            self.Entry_consultar_total.insert(END, col9)
            self.Entry_consultar_quantidade.delete(0, END)

    def corrigir_valor_consulta(self):
        self.consulta_data = self.entry_dataa.get()
        self.consulta_cliente = self.Entry_consultar_cliente.get()
        self.consulta_seguimento = self.Entry_consultar_seguimento.get()
        self.consulta_tema = self.Entry_consultar_tema.get()
        self.consulta_descricao = self.Entry_consultar_descricao.get()
        self.consulta_tamanho = self.Entry_consultar_tamanho.get()
        self.consulta_quantidade = self.Entry_consultar_quantidade.get()
        self.consulta_valor = self.Entry_consultar_valor.get()
        self.sql()
        msg= messagebox.askyesno('REMOVER',f'DESEJA RETIRAR {self.consulta_quantidade} PEÇA(S)\n'
                                           f' DO SEGUIMENTO {self.consulta_seguimento} , DA VENDA FEITA PARA\n '
                                           f'  {self.consulta_cliente} ?')
        if msg == TRUE:
            quary = self.cursor.execute(
                f"""SELECT quantidade_venda from seguimento_venda
                                            WHERE seguimento_venda = '{self.consulta_seguimento}' 
                                            and tema_venda = '{self.consulta_tema}'
                                            and cliente= '{self.consulta_cliente}'""")
            for i in quary:
                qquantidade = i[0]
                if qquantidade >= int(self.consulta_quantidade):
                    self.corrigirr = self.cursor.execute(f"""
                    UPDATE seguimento_venda 
                    SET  quantidade_venda = quantidade_venda - '{self.consulta_quantidade}'                       
                    WHERE cliente='{self.consulta_cliente}'
                    AND   seguimento_venda ='{self.consulta_seguimento}'
                    AND data_venda= '{self.consulta_data}'
                    AND tema_venda = '{self.consulta_tema}'
                    AND descricao_venda= '{self.consulta_descricao}'
                    AND tamanho_venda= '{self.consulta_tamanho}'
                    AND valor_venda= '{self.consulta_valor}'           
                                         """)
                    messagebox.showinfo(
                        '', f'FOI ADICIONADO EM SEU ESTOQUE {self.consulta_quantidade}  UN \n'
                            f'PARA O SEGUIMENTO {self.consulta_seguimento} '
                            f'AQUAL VOCE FEZ ALTERAÇAO NA VENDA DE {self.consultar_cliente}')
                    self.corrigirr = self.cursor.execute(
                        f"""UPDATE seguimento_venda 
                        SET quantidade = quantidade + '{self.consulta_quantidade}'                       
                                                        WHERE   seguimento ='{self.consulta_seguimento}'
                                                        AND tema = '{self.consulta_tema}'
                                                        AND descricao= '{self.consulta_descricao}'
                                                        AND tamanho= '{self.consulta_tamanho}'
                                                        AND valor= '{self.consulta_valor}'           
                                                                                         """)
                    self.cursor.commit()
                    self.desconectar()
                    self.limpa_campo_consulta()

                    messagebox.showinfo('',
                                        f'FOI ADICIONADO NO SEGUIMENTO {self.consulta_seguimento} '
                                        f'A QUANTIDADE DE {self.consulta_quantidade} UN')


                else:
                    messagebox.showerror(
                        'ERRO!!!', 'ESSA AÇAO NAO PODE SER COMPLETADA'
                                   f'  POIS A QUANTIDADE  VENDIDA E DE {qquantidade} PEÇAS(s) '
                                   f' {self.consulta_seguimento}\n'
                                   f'EA QUANTIDADE QUE DESEJA RETIRA E DE {self.consulta_quantidade} PEÇA(s) '
                                   )
                    messagebox.showwarning('AVISA', 'CASO QUEIRA ADICIONA MAIS ITENS VA NA OPÇAO ADICIONAR')

    def adicionar_iten(self):
        self.adiciona_data = self.entry_dataa.get()
        self.adiciona_cliente = self.Entry_consultar_cliente.get()
        self.adiciona_seguimento = self.Entry_consultar_seguimento.get()
        self.adiciona_tema = self.Entry_consultar_tema.get()
        self.adiciona_descricao = self.Entry_consultar_descricao.get()
        self.adiciona_tamanho = self.Entry_consultar_tamanho.get()
        self.adiciona_quantidade = self.Entry_consultar_quantidade.get()
        self.adiciona_valor = self.Entry_consultar_valor.get()

        self.adiciona = messagebox.askyesno(
            '', f'DESEJA ADICIONA {self.adiciona_quantidade} UN(S) DO SEGUIMENTO {self.adiciona_seguimento}')
        try:
            if self.adiciona == TRUE:
                Tabela = self.cursor.execute(
                    f"""SELECT quantidade FROM seguimento_venda 
                                            WHERE seguimento = '{self.adiciona_seguimento}'
                                            AND tamanho = '{self.adiciona_tamanho}'
                                            AND tema = '{self.adiciona_tema}'""")
                for seguimento in Tabela:
                    Estoque = seguimento[0]
                    if Estoque >= int(self.adiciona_quantidade):
                        self.registrar = self.cursor.execute(
                            f"""UPDATE seguimento_venda 
                                SET quantidade_venda = quantidade_venda + '{self.adiciona_quantidade}'
                                WHERE seguimento_venda = '{self.adiciona_seguimento}'
                                AND tema_venda = '{self.adiciona_tema}'
                                AND tamanho_venda = '{self.adiciona_tamanho}'""")
                        self.vender = self.cursor.execute(
                            f"""UPDATE seguimento_venda 
                            SET quantidade = quantidade  - '{self.adiciona_quantidade}'
                            WHERE seguimento =  '{self.adiciona_seguimento}'
                            AND  tema = '{self.adiciona_tema}'
                            AND tamanho = '{self.adiciona_tamanho}'""")
                        messagebox.showinfo(
                            '', f'FORAM ADICIONADAS {self.adiciona_quantidade} PEÇA(S) {self.adiciona_seguimento}\n'
                                f'PARA O CLIENTE {self.adiciona_cliente}')
                        self.cursor.commit()
                        self.limpa_campo_consulta()
                        self.desconectar()

                    else:
                        messagebox.showerror(
                            'ERRO', f'NÃO E POSSIVEL ADICIONAR O SEGUIMENTO PARA {self.adiciona_cliente} \n'
                                    f'SEU ESTOQUE E DE ( {Estoque} ) PEÇA(S)\n'
                                    f'EA QUANTIDADE QUE VOÇÊ DESEJA ADICIONAR E DE\n '
                                    f'{self.adiciona_quantidade} PEÇA(S) '
                                    f' POR FAVOR ADICONE UMA QUANTIDADE  REFERENTE  A SEU ESTOQUE')
        except:
            pass

    def excluir_venda(self):
        self.sql()
        self.consulta_data = self.entry_dataa.get()
        self.consulta_cliente = self.Entry_consultar_cliente.get()
        self.consulta_seguimento = self.Entry_consultar_seguimento.get()
        self.consulta_tema = self.Entry_consultar_tema.get()
        self.consulta_descricao = self.Entry_consultar_descricao.get()
        self.consulta_tamanho = self.Entry_consultar_tamanho.get()
        self.consulta_quantidade = self.Entry_consultar_quantidade.get()
        self.consulta_valor = self.Entry_consultar_valor.get()

        if self.consulta_seguimento == '' \
                or self.consulta_tema == '' \
                or self.consulta_descricao == '' \
                or self.consulta_tamanho == '' \
                or self.consulta_valor == '':
            messagebox.showinfo('', 'E NECESSARIO ANTES DE FAZER A EXCLUSAO,REALIZAR A PESQUISA'
                                    ' PARA INFOMAR CORRETAMENTE OS DADOS QUE SERAO EXCLUIDOS!')
        else:
            self.msgExcluir = messagebox.askyesno(
                '', f'DESEJA EXCLUIR AVENDA DE {self.consulta_seguimento} PARA {self.consultar_cliente}?')
        try:
            if self.msgExcluir == TRUE:
                self.msgExcluir1 = messagebox.askyesno(
                    '', f'AO EXCLUIR A VENDA DO SEGUIMENTO {self.consulta_seguimento} \n'
                        f'REALIZADA PARA {self.consultar_cliente}'
                        f'SERÁ FEITO O EXTORNO DE {self.consulta_quantidade} PEÇA(S) PARA SEU ESTOQUE DESEJA CONTINUAR?'
                )
                if self.msgExcluir1 == TRUE:
                    validar = self.cursor.execute(f"""
                       SELECT quantidade_venda FROM seguimento_venda WHERE cliente= '{self.consulta_cliente}'               
                       AND  seguimento_venda = '{self.consulta_seguimento}'
                       AND tema_venda ='{self.consulta_tema}'
                       AND tamanho_venda ='{self.consulta_tamanho}'
                       AND descricao_venda ='{self.consulta_descricao}'
                       """)
                    for q in validar:
                        excluir = q[0]
                        if int(self.consulta_quantidade) == excluir:
                            self.corrigirr = self.cursor.execute(
                                f""" DELETE seguimento_venda   WHERE cliente='{self.consulta_cliente}'
                                                                AND   seguimento_venda ='{self.consulta_seguimento}'
                                                                AND data_venda= '{self.consulta_data}'
                                                                AND tema_venda = '{self.consulta_tema}'
                                                                AND descricao_venda= '{self.consulta_descricao}'
                                                                AND tamanho_venda= '{self.consulta_tamanho}'
                                                                AND quantidade_venda='{self.consulta_quantidade}'
                                                                AND valor_venda= '{self.consulta_valor}'           
                                                                                                  """)
                            messagebox.showinfo(
                                '', f'FOI ADICIONADO AO SEU SEGUIMENTO {self.consulta_seguimento} '
                                    f'A QUANTIDADE DE {self.consulta_quantidade}')
                            self.corrigirr = self.cursor.execute(
                                f"""UPDATE seguimento_venda SET  quantidade = quantidade + '{self.consulta_quantidade}'                       
                                                            WHERE   seguimento ='{self.consulta_seguimento}'
                                                            AND tema = '{self.consulta_tema}'
                                                            AND descricao= '{self.consulta_descricao}'
                                                            AND tamanho= '{self.consulta_tamanho}'
                                                            AND valor= '{self.consulta_valor}'           
                                                            """)
                            self.cursor.commit()
                            self.desconectar()
                            self.limpa_campo_consulta()

                        else:
                            messagebox.showerror(
                                                'ERRO !', f'PARA EXCLUIR A VENDA VOCÊ DEVERA INFORMA CORRETAMENTE '
                                                    f'A QUANTIDADE VENDIDA,'
                                                    f' VOCÊ ESTA INFORMANDO UMA QUANTIDADE DE\n '
                                                    f' {self.consulta_quantidade} PEÇA(S) '
                                                    f'{self.consulta_seguimento}\n '
                                                    f'PARA UMA VENDA DE  \n {excluir}  PEÇA(S)'
                                                    f' REALIZADA PARA O CLIENTE >>> {self.consulta_cliente}')
        except:
            pass
#                              FUNÇAO DE ESTOQUE MINIMO  ATRIBUIDA AO BOTAO CONFIGURAR ESTOQUE                         #

    def estoque_minimo(self):
        self.app7 = Toplevel()
        self.app7.geometry('970x600+300+56')
        self.app7.resizable(False, False)
        self.app7.transient(self.app)
        self.app7.focus_force()
        self.app7.grab_set()
        self.app7.configure(bg='white')
        self.frameEstoque = Frame(self.app7, bg="white", width=955, height=300, borderwidth=9, relief="raised",
                                  background='#7f7f7f')
        self.frameEstoque.place(x=10, y=10)
        self.frameinferiorestoque = Frame(self.app7, bg="white", width=955, height=270, borderwidth=9,
                                          relief="raised",
                                          background='#7f7f7f')
        self.frameinferiorestoque.place(x=10, y=315)

        self.lb_estoque = Label(self.app7, text='ADICIONAR  ESTOQUE MINIMO ', bg='#574c26',
                                fg='white', font='times 15', relief='raised')
        self.lb_estoque.place(x=350, y=23)

        self.lb_estoque_seguimento = Label(self.app7, text='ID', bg='#574c26',
                                           fg='white', relief='raised', font='times 14', padx=15)
        self.lb_estoque_seguimento.place(x=35, y=60)

        self.lb_estoque_seguimento = Label(self.app7, text='SEGUIMENTO', bg='#574c26',
                                           fg='white', relief='raised', font='times 14')
        self.lb_estoque_seguimento.place(x=150, y=60)

        self.lb_estoque_tema = Label(self.app7, text='TEMA ou COR', bg='#574c26',
                                     fg='white', relief='raised', font='times 14')
        self.lb_estoque_tema.place(x=335, y=60)

        self.lb_estoque_descricao = Label(self.app7, text='DESCRIÇAO', bg='#574c26',
                                          fg='white', relief='raised', font='times 14')
        self.lb_estoque_descricao.place(x=525, y=60)

        self.lb_estoque_tamanho = Label(self.app7, text='TAMANHO', bg='#574c26',
                                        fg='white', relief='raised', font='times 14')
        self.lb_estoque_tamanho.place(x=675, y=60)

        self.lb_estoque_quantidade = Label(self.app7, text=' QTDD MINIMA', bg='#574c26',
                                           fg='white', relief='raised', font='times 14')
        self.lb_estoque_quantidade.place(x=810, y=60)

        self.treview_estoque()

        self.entry_estoque_id = Entry(self.app7, font='times 15', width=7)
        self.entry_estoque_id.place(x=25, y=95)

        self.entry_estoque_seguimento = Entry(self.app7, font='times 15', width=15)
        self.entry_estoque_seguimento.place(x=135, y=95)
        self.entry_estoque_tema = Entry(self.app7, font='times 15', width=15)
        self.entry_estoque_tema.place(x=320, y=95)
        self.entry_estoque_descricao = Entry(self.app7, font='times 15', width=15)
        self.entry_estoque_descricao.place(x=505, y=95)
        self.entry_estoque_tamanho = Entry(self.app7, font='times 15', width=7)
        self.entry_estoque_tamanho.place(x=690, y=95)
        self.entry_estoque_qtdd_minima = Entry(self.app7, font='times 15 ', width=7)
        self.entry_estoque_qtdd_minima.place(x=850, y=95)

        self.botao = Button(self.app7, text='ADICIONAR SELEÇAO ', bg='#f24f00',
                            fg='WHITE',
                            font='times 13 bold',
                            command=self.inserir_estoque)
        self.botao.place(x=370, y=140)

        #                       FUNCAO DE PESQUISAR SEGUIMENTO DO ESTOQUE                                       #

        self.lb_estoque_pesquisa = Label(self.app7, text='PESQUISAR POR SEGUIMENTO ', bg='#574c26',
                                         fg='white', font='times 15', relief='raised').place(x=25, y=170)
        self.entry_lb_seguimento = Label(self.app7, text='SEGUIMENTO', bg='#7f7f7f',
                                         fg='white', relief='raised', font='times 14').place(x=190, y=210)
        self.entry_lb_tema = Label(self.app7, text='TEMA', bg='#7f7f7f',
                                   fg='white', relief='raised', font='times 14').place(x=190, y=242)
        self.entry_lb_descricao = Label(self.app7, text='DESCRICAO', bg='#7f7f7f',
                                        fg='white', relief='raised', font='times 14').place(x=190, y=272)

        self.entry_pesquisar_seguimento = Entry(self.app7, font='times 15', width=15)
        self.entry_pesquisar_seguimento.place(x=25, y=210)
        self.entry_pesquisar_tema = Entry(self.app7, font='times 15', width=9)
        self.entry_pesquisar_tema.place(x=25, y=240)
        self.entry_pesquisar_descricao = Entry(self.app7, font='times 15', width=15)
        self.entry_pesquisar_descricao.place(x=25, y=270)

        self.pesquisar_seguimento = Button(self.app7, text='PESQUISAR',
                                           font='times 11 bold',
                                           bg='#f24f00',
                                           fg='white',
                                           command=self.consulta_seguimento_estoque)
        self.pesquisar_seguimento.place(x=310, y=270)

        self.fechar_estoque = Button(self.app7, text='FECHAR', font='times 11 bold',
                                     bg='#f24f00',
                                     fg='white',
                                     command=self.app7.destroy)
        self.fechar_estoque.place(x=850, y=25)

        self.fechar_adiciona = Button(self.app7, text='ADD ALL', font='times 11 bold',
                                      bg='#f24f00',
                                      fg='white', command=self.estoque_minimo_all)
        self.fechar_adiciona.place(x=850, y=130)

    def treview_estoque(self):
        self.columns = ('id', 'seguimento', 'tema', 'descricao', 'tamanho', 'qtdd_minima')
        self.treeview_estoque = ttk.Treeview(self.app7, columns=self.columns, show='headings')
        self.treeview_estoque.heading('id', text='ID')
        self.treeview_estoque.heading('seguimento', text='Seguimento')
        self.treeview_estoque.heading('descricao', text='Descricao')
        self.treeview_estoque.heading('tema', text='Tema')
        self.treeview_estoque.heading('tamanho', text='Tamanho')
        self.treeview_estoque.heading('qtdd_minima', text='Qtdd_Minima')

        self.treeview_estoque.column('id', width=10)
        self.treeview_estoque.column('seguimento', width=200)
        self.treeview_estoque.column('tema', width=180)
        self.treeview_estoque.column('descricao', width=150)
        self.treeview_estoque.column('tamanho', width=100)
        self.treeview_estoque.column('qtdd_minima', width=30)
        self.treeview_estoque.place(x=20, y=330, width=905, height=240)
        self.scrollmin = Scrollbar(self.app7, orient=VERTICAL, command=self.treeview_estoque.yview)
        self.treeview_estoque.configure(yscrollcommand=self.scrollmin.set)
        self.scrollmin.place(x=930, y=330, width=25, height=238)
        self.treeview_estoque.bind("<Double-1>", self.double_click_estoque)

    def consulta_seguimento_estoque(self):
        self.sql()
        self.treeview_estoque.delete(*self.treeview_estoque.get_children())
        self.pesquisar_estoque = self.entry_pesquisar_seguimento.get()
        self.pesquisa_Tema = self.entry_pesquisar_tema.get()
        self.pesquisa_Descricao = self.entry_pesquisar_descricao.get()

        if self.pesquisar_estoque == '':
            messagebox.showinfo('', 'PREENCHA O CAMPO  CORRETAMENTE ')
        elif self.pesquisar_estoque != '' and self.pesquisa_Tema != '' and self.pesquisa_Descricao != '':

            self.estoque_1 = self.cursor.execute(f"""SELECT id,
                                                            seguimento,
                                                            tema,
                                                            descricao,
                                                            tamanho,
                                                            qtdd_minima 
                                                            FROM seguimento_venda 
                                                            WHERE seguimento like '{self.pesquisar_estoque}%'
                                                            AND tema LIKE '{self.pesquisa_Tema}%' 
                                                            AND descricao LIKE '{self.pesquisa_Descricao}%'
                                                            ORDER BY seguimento,tema,tamanho
                                                                                """)

            for linha in self.estoque_1:
                self.estoque = linha.id, \
                               linha.seguimento, \
                               linha.tema, \
                               linha.descricao, \
                               linha.tamanho, \
                               linha.qtdd_minima
                self.treeview_estoque.insert('', 'end', values=self.estoque)
        elif self.pesquisar_estoque != '' and self.pesquisa_Tema != '':
            self.estoque_2 = self.cursor.execute(f"""SELECT id,
                                                            seguimento,
                                                            tema,
                                                            descricao,
                                                            tamanho,
                                                            qtdd_minima 
                                                            from seguimento_venda 
                                                            where seguimento like '{self.pesquisar_estoque}%' 
                                                            AND tema like '{self.pesquisa_Tema}%'  
                                                            ORDER BY seguimento,tema,tamanho DESC   """)

            for linha in self.estoque_2:
                self.estoque = linha.id, \
                               linha.seguimento, \
                               linha.tema, \
                               linha.descricao, \
                               linha.tamanho, \
                               linha.qtdd_minima
                self.treeview_estoque.insert('', 'end', values=self.estoque)

        else:
            self.pesquisar_configurar_estoque = self.cursor.execute(f"""SELECT id,
                                                                    seguimento,
                                                                    tema,
                                                                    descricao,
                                                                    tamanho,
                                                                    qtdd_minima 
                                                                    from seguimento_venda 
                                                                    where seguimento LIKE '{self.pesquisar_estoque}%'
                                                                     ORDER BY seguimento,tema,tamanho DESC
                                                                                    """)
            for linha in self.cursor.fetchall():
                self.estoque = linha.id, \
                               linha.seguimento, \
                               linha.tema, \
                               linha.descricao, \
                               linha.tamanho, \
                               linha.qtdd_minima
                self.treeview_estoque.insert('', 'end', values=self.estoque)

    def estoque_minimo_all(self):
        self.sql()
        self.estoque_all_seguimento = self.entry_estoque_seguimento.get()
        self.estoque_all_tema = self.entry_estoque_tema.get()
        self.estoque_all_descricao = self.entry_estoque_descricao.get()
        self.estoque_all_tamanho = self.entry_estoque_tamanho.get()
        self.estoque_all_qttd_minima = self.entry_estoque_qtdd_minima.get()
        if self.estoque_all_seguimento != '' or self.estoque_all_tema != '' or self.estoque_all_descricao != ''\
                or self.estoque_all_tamanho != '' or self.estoque_all_qttd_minima != '':
            messagebox.showwarning('AVISO','ESSA FUNÇÃO CONFIGURA SEU TODO SEU ESTOQUE MINIMO ')
            mensagem= messagebox.askyesno(
            'AVISO !', f'AO USAR A FUNÇAO ADD ALL TODO O ESTOQUE MINIMO DO SEGUIMENTO {self.estoque_all_seguimento} \n'
                                              f'SERA CONFIGURADO PARA {self.estoque_all_qttd_minima}'
                                              f' PEÇA(S) DESEJA CONTINUAR ?')
            if mensagem == TRUE:
                self.estoqueminimo_all = self.cursor.execute(
                    f"""UPDATE seguimento_venda SET qtdd_minima  = '{self.estoque_all_qttd_minima}' 
                                                WHERE seguimento ='{self.estoque_all_seguimento}'
                                                """)

                messagebox.showinfo('', 'ESTOQUE MINIMO CONFIGURADO COM SUCESSO')
                self.cursor.commit()

            else:
                messagebox.showinfo('', 'NENHUM SEGUIMENTO FOI CONFIGURADO')

        else:
            messagebox.showwarning(
                'CAMPOS OBRIGATÓRIOS','E NECESSARIO INFORMAR CORRENTAMENTE OS DADOS DO SEGUIMENTO,\n'
                                      'QUE DESEJAR ESTABELECER A QUANTIDADE MINIMA PARA ESTOQUE\n '
                                      'PARA ISSO USE A OPÇÃO PESQUISAR POR SEGUIMENTO')

    def limpa_campo_estoque(self):
        self.entry_estoque_id.delete(0, END)
        self.entry_estoque_seguimento.delete(0, END)
        self.entry_estoque_tema.delete(0, END)
        self.entry_estoque_descricao.delete(0, END)
        self.entry_estoque_tema.delete(0, END)
        self.entry_estoque_tamanho.delete(0, END)
        self.entry_estoque_qtdd_minima.delete(0, END)

    def double_click_estoque(self, event):
        self.treeview_estoque.selection()
        self.limpa_campo_estoque()
        for i in self.treeview_estoque.selection():
            col1, col2, col3, col4, col5, col6 = self.treeview_estoque.item(i, 'values')
            self.entry_estoque_id.insert(END, col1)
            self.entry_estoque_seguimento.insert(END, col2)
            self.entry_estoque_tema.insert(END, col3)
            self.entry_estoque_descricao.insert(END, col4)
            self.entry_estoque_tamanho.insert(END, col5)
            self.entry_estoque_qtdd_minima.insert(END, col6)
            self.entry_estoque_qtdd_minima.delete(0, END)

    def inserir_estoque(self):
        self.sql()
        self.estoque_id = self.entry_estoque_id.get()
        self.estoque_seguimento = self.entry_estoque_seguimento.get()
        self.estoque_tema = self.entry_estoque_tema.get()
        self.estoque_descricao = self.entry_estoque_descricao.get()
        self.estoque_tamanho = self.entry_estoque_tamanho.get()
        self.estoque_qttd_minima = self.entry_estoque_qtdd_minima.get()
        if self.estoque_id == '' \
                or self.estoque_seguimento == '' \
                or self.estoque_tema == '' \
                or self.estoque_descricao == '' \
                or self.estoque_tamanho == '' \
                or self.estoque_qttd_minima == '':
            messagebox.showinfo('', 'REALIZA A BUSCA  DO SEGUIMENTO PARA DEFINIR SEU ESTOQUE MINIMO')
        else:

            self.msgEstoque = messagebox.askyesno('', f'DESEJA ADICIONAR  {self.estoque_qttd_minima} PEÇA(S) '
                                                      f' PARA SEU ESTOQUE MINIMO DE {self.estoque_seguimento} ?')
            if self.msgEstoque == TRUE:
                self.msgEstoque1 = messagebox.askyesno(
                    '', f'SEU ESTOQUE MINIMO PARA O SEGUIMENTO  {self.estoque_seguimento} '
                        f' FICARA COM  {self.estoque_qttd_minima} PEÇA(S) '
                        f'VOCÊ PODERA ALTERA A QUAL QUER MOMENTO  SEU ESTOQUE MINIMO\n '
                        f'DESEJA CONTINUAR ?')

                if self.msgEstoque1 == TRUE:
                    self.estoqueminimo = self.cursor.execute(
                        f"""UPDATE seguimento_venda SET qtdd_minima  = '{self.estoque_qttd_minima}' 
                                                                        WHERE seguimento ='{self.estoque_seguimento}'
                                                                        AND tema= '{self.estoque_tema}'
                                                                        AND descricao = '{self.estoque_descricao}'
                                                                        AND tamanho = '{self.estoque_tamanho}'
                                                                        AND id = '{self.estoque_id}'""")
                    messagebox.showinfo('', 'ESTOQUE MINIMO CONFIGURADO COM SUCESSO')

                    self.limpa_campo_estoque()
                    self.cursor.commit()
                    self.desconectar()
    #                                        FUNÇAO ATRIBUIDA AO MENU      CONSULTA/EXPORTA/ESTOQUE                    #

    def exportar_estoque(self):
        self.app8 = Toplevel()
        self.app8.geometry('970x600+300+56')
        self.app8.resizable(False, False)
        self.app8.transient(self.app)
        self.app8.focus_force()
        self.app8.grab_set()
        self.app8.configure(bg='white')

        self.frameexportar = Frame(self.app8, bg="white", width=955, height=300, borderwidth=9, relief="raised",
                                   background='#7f7f7f')
        self.frameexportar.place(x=10, y=10)

        self.frameinferiorExporta = Frame(self.app8, bg="white", width=955, height=270, borderwidth=9,
                                          relief="raised",
                                          background='#7f7f7f')
        self.frameinferiorExporta.place(x=10, y=315)

        self.lb_exporta = Label(self.app8, text='CONSULTA DETALHADA DE ESTOQUE', bg='#574c26',
                                fg='white', font='times 15', relief='raised')
        self.lb_exporta.place(x=320, y=23)

        # self.lb_exporta_seguimento = Label(self.app8, text='ID', bg='#574c26',
        #                                    fg='white', relief='raised', font='times 14', padx=15)
        # self.lb_exporta_seguimento.place(x=110, y=95)

        self.lb_exporta_seguimento = Label(self.app8, text='SEGUIMENTO', bg='#574c26',
                                           fg='white', relief='raised', font='times 14')
        self.lb_exporta_seguimento.place(x=190, y=130)

        self.lb_exporta_tema = Label(self.app8, text='TEMA ou COR', bg='#574c26',
                                     fg='white', relief='raised', font='times 14')
        self.lb_exporta_tema.place(x=190, y=170)

        self.lb_exporta_quantidade = Label(self.app8, text=' QTDD MINIMA', bg='#574c26',
                                           fg='white', relief='raised', font='times 14')
        self.lb_exporta_quantidade.place(x=420, y=100)

        # self.exporta_estoque_id = Entry(self.app8, font='times 15', width=7)
        # self.exporta_estoque_id.place(x=25, y=95)
        self.entry_exporta_seguimento = Entry(self.app8, font='times 15', width=15)
        self.entry_exporta_seguimento.place(x=25, y=130)
        self.entry_exporta_tema = Entry(self.app8, font='times 15', width=15)
        self.entry_exporta_tema.place(x=25, y=170)
        self.entry_exporta_qtdd_minima = Entry(self.app8, font='times 15 ', width=7)
        self.entry_exporta_qtdd_minima.place(x=450, y=130)
        self.treview_exporta()

#                   BOTAO EXPORTAR TRAS O VALOR DO ESTOQUE MINIMO CADASTRADO E INFERIOR NO BANCO DO DADOS              #
        self.estque=PhotoImage(file='imagens/busca.png')
        self.botaoExporta = Button(self.app8, text='ESTOQUE MINIMO', bg='WHITE',
                                   image=self.estque,
                                   compound=TOP,
                                   fg='#023769',
                                   font='times 13 bold',
                                   command=self.balanco_estoque)
        self.botaoExporta.place(x=780, y=140)

#                                FUNÇAO ATRIBUIDA PARA EXPORTA O ESTOQUE PARA EXEL                                   #
        self.exell = PhotoImage(file='imagens/exel.png')
        self.botaoExcel = Button(self.app8, text='EXPORTA ESTOQUE',
                                 bg='WHITE',
                                 image=self.exell,
                                 compound=TOP,
                                 fg='GREEN',
                                 font='times 13 bold',
                                 command=self.exporta_balanco)
        self.botaoExcel.place(x=400, y=220)

#                        USA OS ENTRYS ACIMA DO BOTAO PARA COLETAR AS INFORMAÇOES INSERIDAS DO USUARIO                #
        self.pesquisaDetalhada= PhotoImage(file='imagens/pesquisadetalhada.png')
        self.ex_seguimento = Button(self.app8, text='DETALHADA',
                                    image=self.pesquisaDetalhada,
                                    compound=TOP,
                                    font='times 11 bold',
                                    bg='#f24f00',
                                    fg='white',
                                    command=self.balanco_detalhado)
        self.ex_seguimento.place(x=30, y=210)

#                                BOTAO FECHA A JANELA                                                                 #

        self.close_estoque = Button(self.app8, text='FECHAR', font='times 11 bold',
                                    bg='#f24f00',
                                    fg='white',
                                    command=self.app8.destroy)
        self.close_estoque.place(x=850, y=25)

    def treview_exporta(self):
        self.columns = ('id', 'seguimento', 'tema', 'descricao', 'tamanho', 'quantidade', 'qtdd_minima')
        self.treeview_exporta = ttk.Treeview(self.app8, columns=self.columns, show='headings')
        self.treeview_exporta.heading('id', text='ID')
        self.treeview_exporta.heading('seguimento', text='Seguimento')
        self.treeview_exporta.heading('descricao', text='Descricao')
        self.treeview_exporta.heading('tema', text='Tema')
        self.treeview_exporta.heading('tamanho', text='Tamanho')
        self.treeview_exporta.heading('quantidade', text='Quantidade')
        self.treeview_exporta.heading('qtdd_minima', text='Qtdd_Minima')

        self.treeview_exporta.column('id', width=10)
        self.treeview_exporta.column('seguimento', width=200)
        self.treeview_exporta.column('tema', width=180)
        self.treeview_exporta.column('descricao', width=150)
        self.treeview_exporta.column('tamanho', width=100)
        self.treeview_exporta.column('quantidade', width=100)
        self.treeview_exporta.column('qtdd_minima', width=60)
        self.scrollexp = Scrollbar(self.app8, orient=VERTICAL, command=self.treeview_exporta.yview)
        self.treeview_exporta.configure(yscrollcommand=self.scrollexp.set)
        self.scrollexp.place(x=930, y=330, width=25, height=238)
        self.treeview_exporta.place(x=20, y=330, width=905, height=240)

    def balanco_estoque(self):
        self.sql()
        self.treeview_exporta.delete(*self.treeview_exporta.get_children())
        self.balanco = self.cursor.execute(f"""SELECT id,seguimento,tema,descricao,tamanho,quantidade,qtdd_minima 
        FROM seguimento_venda WHERE quantidade  <= qtdd_minima ORDER BY seguimento,tema,tamanho                              
                                                    """)

        for linha in self.balanco:
            self.balanco_Estoque = linha.id, \
                                   linha.seguimento, \
                                   linha.tema, \
                                   linha.descricao, \
                                   linha.tamanho, \
                                   linha.quantidade, \
                                   linha.qtdd_minima
            self.treeview_exporta.insert('', 'end', values=self.balanco_Estoque)

        self.desconectar()

    def balanco_detalhado(self):
        self.sql()

        self.treeview_exporta.delete(*self.treeview_exporta.get_children())
        self.exp_seguimento = self.entry_exporta_seguimento.get()
        self.exp_tema = self.entry_exporta_tema.get()
        self.exp_quantidade = self.entry_exporta_qtdd_minima.get()

        if self.exp_seguimento != '' and self.exp_quantidade != '':
            self.consulta_estoque = self.cursor.execute(
                f"""SELECT id,seguimento,tema,descricao,tamanho,quantidade,qtdd_minima
                    FROM seguimento_venda WHERE seguimento LIKE '{self.exp_seguimento}%'
                    AND quantidade <= '{self.exp_quantidade}'
                    ORDER BY seguimento,tema,tamanho""")
            for linha in self.cursor.fetchall():
                self.resultados = linha.id, linha.seguimento, linha.tema, linha.descricao, linha.tamanho,\
                                  linha.quantidade, linha.qtdd_minima
                self.treeview_exporta.insert('', 'end', values=self.resultados)


        elif self.exp_tema != '' and self.exp_quantidade != '':
                self.consulta_estoque1 = self.cursor.execute(
                    f"""SELECT id,seguimento,tema,descricao,tamanho,quantidade,qtdd_minima
                        FROM seguimento_venda WHERE tema = '{self.exp_tema}'
                        AND quantidade <= '{self.exp_quantidade}'
                        ORDER BY seguimento,tema,tamanho""")

                for linha in self.consulta_estoque1:
                    self.resultados = linha.id, linha.seguimento, linha.tema, linha.descricao, linha.tamanho, \
                                      linha.quantidade, linha.qtdd_minima
                    self.treeview_exporta.insert('', 'end', values=self.resultados)

        elif self.exp_quantidade != '':
            self.consulta_estoque1 = self.cursor.execute(
                f"""SELECT * FROM seguimento_venda WHERE quantidade <= '{self.exp_quantidade}'
                               ORDER BY seguimento,tema,tamanho""")

            for linha in self.consulta_estoque1:
                self.resultados = linha.id, linha.seguimento, linha.tema, linha.descricao, linha.tamanho, \
                                  linha.quantidade, linha.qtdd_minima
                self.treeview_exporta.insert('', 'end', values=self.resultados)
        else:
            messagebox.showwarning(
                'AVISO!','PARA REALIZAR ALGUMAS CONSULTA  VOCÊ DEVERÁ INFORMA A CONSULTA QUE DESEJA\n '
                                'VOCÊ PODE ESCOLHER O SEGUIMENTO EA QUANTIDADE MINIMA QUE DESEJA \n '
                                'EXE: SEGUIMENTO X QUANTIDADE \n  (CONSULTARA A QUANTIDADE MINIMA DO SEGUIMENTO )\n'
                                'TEMA X QUANTIDADE \n  ( CONSULTARA A QUANTIDADE MINIMA DO TEMA )\n '
                                'X QUANTIDADE \n ( CONSULTATA TODOS OS SEGUIEMENTOS E TEMAS APARTIR\n '
                                'DA QUANTIDADE MINIMA INFORMADA )' )

    def exporta_balanco(self):

        import pandas as pd
        import pyodbc

        self.exp_seguimento = self.entry_exporta_seguimento.get()
        self.exp_tema = self.entry_exporta_tema.get()
        self.exp_quantidade = self.entry_exporta_qtdd_minima.get()

        if self.exp_seguimento != '' and self.exp_quantidade != '':
            expor= messagebox.askyesno('', f'SERÁ EXPORTADO TODOS OS SEGUIMENTOS REFERENTE A {self.exp_seguimento} \n'
                                           f'COM QUANTIDADE IGUAL OU MENOR QUE {self.exp_quantidade} PEÇA(S) '
                                           f'DESEJA CONTINUAR?')
            if expor== TRUE:
                self.dados_conexao = ("Driver={SQL Server};"
                                      "Server=DESKTOP-RFPLHG0;"
                                      "Database=seguimentos_cadastrados;")
                self.conexao = pyodbc.connect(self.dados_conexao)
                self.cursor = self.conexao.cursor()

                self.exp_balanco = pd.read_sql(
                    f"SELECT seguimento,tema,descricao,tamanho,quantidade,valor FROM "
                    f"seguimento_venda  WHERE seguimento like '{self.exp_seguimento}%'"
                    f"AND quantidade  <= '{self.exp_quantidade}'", self.conexao)
                self.balanco = pd.DataFrame(data=self.exp_balanco)
                self.seguimento_expp = self.exp_balanco[
                    ['seguimento','tema', 'descricao', 'tamanho', 'quantidade', 'valor']] \
                    .groupby(['seguimento','tema', 'descricao', 'tamanho']).sum()
                self.seguimento_expp.to_excel('seguimentoDeBlanço.xlsx')
                messagebox.showinfo('', 'EXPORTAÇAO BEM SUCEDIDA')
            else:
                messagebox.showinfo('','NENHUM VALOR FOI EXPORTADO')


        elif self.exp_tema != '' and self.exp_quantidade != '' and self.exp_seguimento != '':
            menss= messagebox.askyesno('', f'SERÁ EXPORTADO TODOS OS SEGUIMENTOS REFERENTE AO TEMA {self.exp_tema} \n'
                                           f'COM QUANTIDADE IGUAL OU MENOR QUE {self.exp_quantidade} PEÇA(S) '
                                           f'DESEJA CONTINUAR?')
            if menss == TRUE:
                self.dados_conexao = ("Driver={SQL Server};"
                                      "Server=DESKTOP-RFPLHG0;"
                                      "Database=seguimentos_cadastrados;")
                self.conexao = pyodbc.connect(self.dados_conexao)
                self.cursor = self.conexao.cursor()

                self.exp_balanco = pd.read_sql(
                    f"SELECT seguimento,tema,descricao,tamanho,quantidade,valor" 
                    f" FROM seguimento_venda  "
                    f"WHERE tema like '{self.exp_tema}%'"
                    f"AND seguimento = '{self.exp_seguimento}'"
                    f"AND quantidade  <= '{self.exp_quantidade}'", self.conexao)

                self.balanco = pd.DataFrame(data=self.exp_balanco)
                self.seguimento_expp = self.exp_balanco[
                    ['seguimento', 'tema', 'descricao', 'tamanho', 'quantidade', 'valor']] \
                    .groupby(['seguimento', 'tema', 'descricao', 'tamanho']).sum()
                self.seguimento_expp.to_excel('seguimentotemaDeBalanço.xlsx')

                messagebox.showinfo('', 'EXPORTAÇAO BEM SUCEDIDA')
            else:
                messagebox.showinfo('','NENHUM VALOR FOI EXPORTADO')




        elif self.exp_tema != '' and self.exp_quantidade != '':
            menss= messagebox.askyesno('', f'SERÁ EXPORTADO TODOS OS SEGUIMENTOS REFERENTE AO TEMA {self.exp_tema} \n'
                                           f'COM QUANTIDADE IGUAL OU MENOR QUE {self.exp_quantidade} PEÇA(S) '
                                           f'DESEJA CONTINUAR?')
            if menss == TRUE:
                self.dados_conexao = ("Driver={SQL Server};"
                                      "Server=DESKTOP-RFPLHG0;"
                                      "Database=seguimentos_cadastrados;")
                self.conexao = pyodbc.connect(self.dados_conexao)
                self.cursor = self.conexao.cursor()

                self.exp_balanco = pd.read_sql(
                    f"SELECT seguimento,tema,descricao,tamanho,quantidade,valor" 
                    f" FROM seguimento_venda  "
                    f"WHERE tema like '{self.exp_tema}%'"
                    f"AND quantidade  <= '{self.exp_quantidade}'", self.conexao)

                self.balanco = pd.DataFrame(data=self.exp_balanco)
                self.seguimento_expp = self.exp_balanco[
                    ['seguimento', 'tema', 'descricao', 'tamanho', 'quantidade', 'valor']] \
                    .groupby(['seguimento', 'tema', 'descricao', 'tamanho']).sum()
                self.seguimento_expp.to_excel('TemaDeBalanço.xlsx')

                messagebox.showinfo('', 'EXPORTAÇAO BEM SUCEDIDA')
            else:
                messagebox.showinfo('','NENHUM VALOR FOI EXPORTADO')

        elif self.exp_quantidade != '':
            mens= messagebox.askyesno('', f'SERÁ EXPORTADO TODOS OS SEGUIMENTOS COM BASE NA QUANTIDADE ESTABELECIDA.\n '
                                          f' QUE SERÁ IGUAL OU MENOR QUE {self.exp_quantidade} PEÇA(S)\n ' 
                                           f'DESEJA CONTINUAR ?')
            if mens == TRUE:
                self.dados_conexao = ("Driver={SQL Server};"
                                      "Server=DESKTOP-RFPLHG0;"
                                      "Database=seguimentos_cadastrados;")
                self.conexao = pyodbc.connect(self.dados_conexao)
                self.cursor = self.conexao.cursor()

                self.exp_balanco = pd.read_sql(
                    f"SELECT seguimento,tema,descricao,tamanho,quantidade,valor"
                    f" FROM seguimento_venda "
                    f" WHERE quantidade <= '{self.exp_quantidade}'", self.conexao)

                self.balanco = pd.DataFrame(data=self.exp_balanco)
                self.seguimento_expp = self.exp_balanco[
                    ['seguimento', 'tema', 'descricao', 'tamanho', 'quantidade', 'valor']] \
                    .groupby(['seguimento', 'tema', 'descricao', 'tamanho']).sum()
                self.seguimento_expp.to_excel('QuantidadeDeBalanço.xlsx')

                messagebox.showinfo('', 'EXPORTAÇAO BEM SUCEDIDA')
            else:
                messagebox.showinfo('', 'NENHUM VALOR FOI EXPORTADO')
        else:
            messagebox.showwarning(
                'AVISO!', 'PARA REALIZAR OPÇAO EXPORTA BALANÇO VOCÊ DEVERÁ INFORMA A EXPORTAÇAO QUE DESEJA.\n '
                          'VOCÊ PODE ESCOLHER O SEGUIMENTO EA QUANTIDADE MINIMA QUE DESEJA \n '
                          'EXE: SEGUIMENTO X QUANTIDADE \n  (EXPORTARÁ A QUANTIDADE MINIMA DO SEGUIMENTO )\n'
                          'TEMA X QUANTIDADE \n  ( EXPORTARÁ A QUANTIDADE MINIMA DO TEMA )\n '
                          'X QUANTIDADE \n ( EXPORTARÁ TODOS OS SEGUIEMENTOS E TEMAS APARTIR\n '
                          'DA QUANTIDADE MINIMA INFORMADA )')

    #                                   FUNÇAO DOS MENUS DE EXPORTA VENDAS

    def exporta_vendas(self):
        self.app9 = Toplevel()
        self.app9.geometry('1010x600+333+56')
        self.app9.resizable(False, False)
        self.app9.transient(self.app)
        self.app9.focus_force()
        self.app9.grab_set()
        self.app9.configure(bg='#2a2a29')

        self.frame_exporta_vendas1 = Frame(self.app9, bg="white", width=995, height=150, borderwidth=9, relief="raised",
                                           background='#7f7f7f')
        self.frame_exporta_vendas1.place(x=10, y=10)

        self.frame_inferior = Frame(self.app9, bg="white", width=995, height=400, borderwidth=9, relief="raised",
                                    background='#7f7f7f')
        self.frame_inferior.place(x=10, y=180)

        self.lab_exp_vendas = Label(self.app9, text=' R E S U L T A D O    D A    P E S Q U I S A  ', bg='#7f7f7f',
                                    fg='white',
                                    relief=FLAT,
                                    font='times 20 bold')
        self.lab_exp_vendas.place(x=150, y=190)

        self.treeview_exp_vendas_valores()

#                                           FRAME E BOTAO DE PESQUISA DA FUNÇAO DELETAR                                #

        self.lab_exp_vendas = Label(self.app9, text='BUSCA POR CLIENTE', bg='#6b6a68', fg='white', relief='raised',
                                    font='times 14')
        self.lab_exp_vendas.place(x=630, y=50)

        self.entry_exp_venda_cliente = Entry(self.app9, font='times 12')
        self.entry_exp_venda_cliente.place(x=630, y=90, width=155)


        self.exp_data_venda = Label(self.app9, text='DATA DA VENDA', bg='#7f7f7f', fg='white',
                                    font='times 15 bold')
        self.exp_data_venda.place(x=320, y=50)

        self.exp_entry_data = Entry(self.app9, font='times 15 bold', width=10)
        self.exp_entry_data.place(x=370, y=100)

        self.bt_exp_vendas = Button(self.app9, text='BUSCAR', bg='#e76e0c', fg='white', font='times 10',
                                    command=self.exp_buscar_cliente_sql)
        self.bt_exp_vendas.place(x=790,y=90)

        self.bt_exp_vendas_voltar = Button(self.app9, text='VOLTAR',
                                           command=self.app9.destroy, bg='#7f7f7f', fg='white', relief=FLAT, padx=10,
                                           font='times 12')
        self.bt_exp_vendas_voltar.place(x=730, y=16)

        self.exp_img_calendario = PhotoImage(file='imagens/calendario.png')
        self.exp_bt_calendario = Button(self.app9, text='data', command=self.exp_calendario_vendas,
                                        image=self.exp_img_calendario, bg='#7f7f7f', relief=FLAT)
        self.exp_bt_calendario.place(x=470, y=85)

        self.bt_exp_vendas_voltar = Button(self.app9, text='EXPORTA VENDA',
                                           command=self.expor_venda, bg='#ff661c', fg='white', pady=5,
                                           font='times 15')
        self.bt_exp_vendas_voltar.place(x=25, y=25)

        self.bt_exp_selecao = Button(self.app9, text='EXPORTA SELEÇAO',
                                      command=self.expor_selecao, bg='#ff661c', fg='white', pady=5,
                                      font='times 15')
        self.bt_exp_selecao.place(x=25, y=90)

        self.bt_exp_agrupamento = Button(self.app9, text='EXPORTA AGRUPAMENTO',
                                      command=self.exporta_agrupamento, bg='#ff661c', fg='white', pady=3,
                                      font='times 10')
        self.bt_exp_agrupamento.place(x=615, y=120)

    def exp_calendario_vendas(self):
        self.exp_calendario1 = Calendar(self.app9, fg='gray75', bg='blue', font='times 9 bold', locale='pt_br')
        self.exp_calendario1.place(x=350, y=90)
        self.exp_data_venda = Button(self.app9, text='INSERIR DATA', bg='#021f53', fg='white',
                                     font='times 10 bold', command=self.exp_inserir_data_venda)
        self.exp_data_venda.place(x=400, y=275)

    def exp_inserir_data_venda(self):
        dataini = self.exp_calendario1.get_date()
        self.exp_calendario1.destroy()
        self.exp_entry_data.delete(0, END)
        self.exp_entry_data.insert(END, dataini)
        self.exp_data_venda.destroy()

    def treeview_exp_vendas_valores(self):
        self.columns = (
            'data', 'cliente', 'seguimento', 'tema', 'descricao', 'tamanho', 'quantidade', 'valor', 'total_venda')
        self.treeview_exp_vendas = ttk.Treeview(self.app9, columns=self.columns, show='headings')
        self.treeview_exp_vendas.heading('data', text='Data')
        self.treeview_exp_vendas.heading('cliente', text='Cliente')
        self.treeview_exp_vendas.heading('seguimento', text='Seguimento')
        self.treeview_exp_vendas.heading('tema', text='Tema')
        self.treeview_exp_vendas.heading('descricao', text='Descriçao')
        self.treeview_exp_vendas.heading('tamanho', text='Tamanho')
        self.treeview_exp_vendas.heading('quantidade', text='QTDD')
        self.treeview_exp_vendas.heading('valor', text='Valor')
        self.treeview_exp_vendas.heading('total_venda', text='Total')

        self.treeview_exp_vendas.column('cliente', width=110)
        self.treeview_exp_vendas.column('data', width=75)
        self.treeview_exp_vendas.column('seguimento', width=200)
        self.treeview_exp_vendas.column('tema', width=160)
        self.treeview_exp_vendas.column('descricao', width=140)
        self.treeview_exp_vendas.column('tamanho', width=60)
        self.treeview_exp_vendas.column('quantidade', width=50)
        self.treeview_exp_vendas.column('valor', width=50)
        self.treeview_exp_vendas.column('total_venda', width=60)
        self.scroll_exp_venda = Scrollbar(self.app9, orient=VERTICAL, command=self.treeview_exp_vendas.yview)
        self.treeview_exp_vendas.configure(yscrollcommand=self.scroll_exp_venda.set)
        self.scroll_exp_venda.place(x=970, y=230, width=25, height=328)
        self.treeview_exp_vendas.place(x=20, y=230, width=943, height=330)
        self.treeview_exp_vendas.bind("<Double-1>", self.double_click_exp)

    def exp_buscar_cliente_sql(self):
        self.sql()
        self.treeview_exp_vendas.delete(*self.treeview_exp_vendas.get_children())
        self.exp_vendas_data = self.exp_entry_data.get()
        self.exp_busc_cliente = self.entry_exp_venda_cliente.get()

        if self.exp_busc_cliente !=''  and self.exp_vendas_data !='':
            self.exp_pes_cliente = self.cursor.execute(
                f""" SELECT * FROM seguimento_venda
                                       WHERE cliente = '{self.exp_busc_cliente}'
                                       AND data_venda= '{self.exp_vendas_data}'
                                       """)
            for linha in self.cursor.fetchall():

                self.exp_pesqui =linha.data_venda, \
                                   linha.cliente, \
                                   linha.seguimento_venda, \
                                   linha.tema_venda, \
                                   linha.descricao_venda, \
                                   linha.tamanho_venda, \
                                   linha.quantidade_venda, \
                                   linha.valor_venda, \
                                   linha.total_venda
                self.treeview_exp_vendas.insert('', END, values=self.exp_pesqui)
        elif self.exp_busc_cliente !='':
            resul= self.cursor.execute(
                f"""SELECT DISTINCT data_venda,cliente
                    FROM seguimento_venda WHERE cliente ='{self.exp_busc_cliente}'""")
            for linha in resul:
                valores=linha.data_venda,linha.cliente
                self.treeview_exp_vendas.insert('', END, values=valores)

        else:
            self.exp_pes_client1e = self.cursor.execute(f""" SELECT DISTINCT data_venda, cliente
                                                    FROM seguimento_venda
                                                      WHERE cliente IS NOT NULL;""")

            for linha in self.cursor.fetchall():
                self.exp_pesqui = linha.data_venda,linha.cliente
                self.treeview_exp_vendas.insert('', 'end', values=self.exp_pesqui)

    def double_click_exp(self, event):
        self.treeview_exp_vendas.selection()
        self.exp_entry_data.delete(0, END)
        self.entry_exp_venda_cliente.delete(0, END)

        for i in self.treeview_exp_vendas.selection():
            col1, col2 = self.treeview_exp_vendas.item(i, 'values')
            self.exp_entry_data.insert(END, col1)
            self.entry_exp_venda_cliente.insert(END, col2)

  #                                         FUNCAO DA REPOSIÇAO DOS MENUS                                              #

    def reposicao(self):
        self.app10 = Toplevel()
        self.app10.geometry('930x580+300+56')
        self.app10.resizable(False, False)
        self.app10.transient(self.app)
        self.app10.focus_force()
        self.app10.grab_set()
        self.app10.configure(bg='#2a2a29')

        self.frame_reposicao = Frame(self.app10, bg="white", width=915, height=290, borderwidth=9, relief="raised",
                                     background='#7f7f7f').place(x=10, y=10)

        self.frame_reposicao_inferior = Frame(self.app10, bg="white", width=915, height=260, borderwidth=9,
                                              relief="raised",
                                              background='#7f7f7f').place(x=10, y=310)

        self.lb_reposicao = Label(self.app10, text='R E P O S I Ç Ã O    D E    E S T O Q U E  ', bg='#7f7f7f',
                                  fg='white', font='times 20', relief=FLAT).place(x=150, y=23)

        self.lb_reposicao_seguimento = Label(self.app10, text='SEGUIMENTO', bg='#7f7f7f',
                                             fg='white', font='times 14', relief='raised').place(x=290, y=60)

        self.lb_reposicao_tema = Label(self.app10, text='TEMA', font='times 14', bg='#7f7f7f',
                                       fg='white', relief='raised').place(x=290, y=100)

        self.lb_reposicao_descricao = Label(self.app10, text='DESCRIÇAO', bg='#7f7f7f',
                                            fg='white', relief='raised', font='times 14').place(x=290, y=140)

        self.lb_reposicao_tamanho = Label(self.app10, text='TAMANHO', bg='#7f7f7f',
                                          fg='white', relief='raised', font='times 14').place(x=120, y=175)

        self.lb_reposicao_quantidade = Label(self.app10, text='QUANTIDADE', bg='#7f7f7f',
                                             fg='white', relief='raised', font='times 14').place(x=120, y=212)

        self.lb_reposicao_valor = Label(self.app10, text='VALOR', bg='#7f7f7f',
                                        fg='white', relief='raised', font='times 14').place(x=120, y=252)

        self.lb_reposicao_id = Label(self.app10, text='ID', bg='#7f7f7f',
                                     fg='white', relief=FLAT, font='times 14 bold').place(x=475, y=60)
        self.treeview_reposicao1()

        # MEUS ENTRYS DA FUNÇAO
        self.entry_reposicao_seguimento = Entry(self.app10, font='times 15', width=25)
        self.entry_reposicao_seguimento.place(x=30, y=60)
        self.entry_reposicao_tema = Entry(self.app10, font='times 15', width=25)
        self.entry_reposicao_tema.place(x=30, y=100)
        self.entry_reposicao_descricao = Entry(self.app10, font='times 15', width=25)
        self.entry_reposicao_descricao.place(x=30, y=140)
        self.entry_reposicao_tamanho = Entry(self.app10, font='times 15', width=7)
        self.entry_reposicao_tamanho.place(x=30, y=175)
        self.entry_reposicao_quantidade = Entry(self.app10, font='times 15 ', width=7)
        self.entry_reposicao_quantidade.place(x=30, y=213)
        self.entry_reposicao_valor = Entry(self.app10, font='times 15 ', width=7)
        self.entry_reposicao_valor.place(x=30, y=253)
        self.entry_reposicao_id = Entry(self.app10, font='times 15 bold ', width=5, bg='#7f7f7f', fg='white',
                                        relief=FLAT)
        self.entry_reposicao_id.place(x=420, y=60)
#                         FRAME E BOTAO DE PESQUISA DA FUNÇAO DELETAR                                     #

        self.rep_cadastro = Label(self.app10, text='BUSCA POR SEGUIMENTO', bg='#6b6a68', fg='white', relief='raised',
                                  font='times 15').place(x=550, y=70)
        self.entry_rep_pesquisa = Entry(self.app10, font='times 15', width=25)
        self.entry_rep_pesquisa.place(x=550, y=105)

        self.reposicao_vendas = Button(self.app10, text='BUSCAR', bg='green', fg='white', font='times 12',
                                       command=self.pesquisar_reposicao)
        self.reposicao_vendas.place(x=710, y=140)

        self.bt_rep_voltar = Button(self.app10, text='VOLTAR', command=self.app10.destroy, bg='#7f7f7f',
                                    fg='orange',
                                    relief=FLAT, pady=5,
                                    font='times 15').place(x=750, y=20)
        self.bt_rep_adicionar = Button(self.app10, text='ADICIONAR',
                                       command=self.adicionar_estoque_reposicao, bg='#574c26', fg='white',
                                       font='times 15').place(x=430, y=245)

        self.bt_rep_retirar = Button(self.app10, text='RETIRAR',
                                     command=self.retirar_excesso, bg='#574c26', fg='white',
                                     font='times 15').place(x=280, y=245)

    def treeview_reposicao1(self):
        self.columns = ('id', 'seguimento', 'tema', 'descricao', 'tamanho', 'quantidade', 'valor')
        self.treeview_reposicao = ttk.Treeview(self.app10, columns=self.columns, show='headings')
        self.treeview_reposicao.heading('id', text='ID')
        self.treeview_reposicao.heading('seguimento', text='Seguimento')
        self.treeview_reposicao.heading('tema', text='Tema')
        self.treeview_reposicao.heading('tamanho', text='Tamanho')
        self.treeview_reposicao.heading('descricao', text='Descriçao')
        self.treeview_reposicao.heading('quantidade', text='Quantidade')
        self.treeview_reposicao.heading('valor', text='Valor')

        self.treeview_reposicao.column('id', width=40)
        self.treeview_reposicao.column('seguimento', width=200)
        self.treeview_reposicao.column('tema', width=180)
        self.treeview_reposicao.column('tamanho', width=50)
        self.treeview_reposicao.column('descricao', width=50)
        self.treeview_reposicao.column('quantidade', width=50)
        self.treeview_reposicao.column('valor', width=50)
        self.treeview_reposicao.place(x=20, y=320, width=870, height=240)
        self.scrollrep = Scrollbar(self.app10, orient=VERTICAL, command=self.treeview_reposicao.yview)
        self.treeview_reposicao.configure(yscrollcommand=self.scrollrep.set)
        self.scrollrep.place(x=895, y=320, width=25, height=240)
        self.treeview_reposicao.bind("<Double-1>", self.double_click_reposicao)

    def pesquisar_reposicao(self):
        self.treeview_reposicao.delete(*self.treeview_reposicao.get_children())
        self.sql()
        self.rep_busca = self.entry_rep_pesquisa.get()

        if self.rep_busca != '':
            self.rep_pesquisar2 = self.cursor.execute(f"""SELECT id,seguimento,tema,descricao,tamanho,quantidade,valor
                                                               from seguimento_venda
                                                                where seguimento LIKE '{self.rep_busca}%'
                                                            ORDER BY tamanho DESC""")
            for linha in self.cursor.fetchall():
                self.i = linha.id, linha.seguimento, linha.tema, linha.descricao, linha.tamanho, linha.quantidade,\
                         linha.valor
                self.treeview_reposicao.insert('', 'end', values=self.i)
        else:
            messagebox.showinfo('', 'DIGITE O SEGUIMENTO QUE DESEJA FAZER A REPOSIÇAO')
        self.desconectar()

    def double_click_reposicao(self, event):
        self.treeview_reposicao.selection()
        self.limpar_tela_treeview_resposicao()
        for y in self.treeview_reposicao.selection():
            col1, col2, col3, col4, col5, col6, col7 = self.treeview_reposicao.item(y, 'values')
            self.entry_reposicao_id.insert(END, col1)
            self.entry_reposicao_seguimento.insert(END, col2)
            self.entry_reposicao_tema.insert(END, col3)
            self.entry_reposicao_descricao.insert(END, col4)
            self.entry_reposicao_tamanho.insert(END, col5)
            self.entry_reposicao_quantidade.insert(END, col6)
            self.entry_reposicao_valor.insert(END, col7)

    def limpar_tela_treeview_resposicao(self):
        self.entry_reposicao_id.delete(0, END)
        self.entry_reposicao_seguimento.delete(0, END)
        self.entry_reposicao_tema.delete(0, END)
        self.entry_reposicao_descricao.delete(0, END)
        self.entry_reposicao_tamanho.delete(0, END)
        self.entry_reposicao_quantidade.delete(0, END)
        self.entry_reposicao_valor.delete(0, END)

    def adicionar_estoque_reposicao(self):
        self.sql()
        self.id = self.entry_reposicao_id.get()
        self.rep_seguimento = self.entry_reposicao_seguimento.get()
        self.rep_tema = self.entry_reposicao_tema.get()
        self.rep_descri = self.entry_reposicao_descricao.get()
        self.rep_tamanho = self.entry_reposicao_tamanho.get()
        self.rep_quantidade = self.entry_reposicao_quantidade.get()

        self.msg_rep = messagebox.askyesno('', 'DESEJA FAZER A REPOSICAO DO SEU ESTOQUE?')
        if self.msg_rep == TRUE:
            self.msg_rep1 = messagebox.askyesno('', f'SERA ADICIONADO A SEU ESTOQUE  {self.rep_quantidade} UNIDADE '
                                                    f'\n PARA O SEGUIMENTO {self.rep_seguimento}')
            if self.msg_rep1 == TRUE:
                self.adicionar_rep = self.cursor.execute(f"""UPDATE  seguimento_venda  SET 
                                                               quantidade = quantidade + {self.rep_quantidade}
                                                              WHERE id = '{self.id}'
                                                              AND seguimento = '{self.rep_seguimento}'
                                                              AND tema = '{self.rep_tema}'
                                                              AND descricao = '{self.rep_descri}'
                                                              AND tamanho = '{self.rep_tamanho}'
                                                                                   """)
                self.cursor.commit()
                self.desconectar()
                messagebox.showinfo('', 'REPOSIÇAO ADICIONADA COM SUCESSO')
            else:
                messagebox.showinfo('', 'NENHUMA REPOSIÇAO FOI REALIZADA')
        else:
            messagebox.showinfo('', 'NENHUMA REPOSIÇAO FOI REALIZADA')

    def retirar_excesso(self):
        self.sql()
        self.id = self.entry_reposicao_id.get()
        self.ret_seguimento = self.entry_reposicao_seguimento.get()
        self.ret_tema = self.entry_reposicao_tema.get()
        self.ret_descri = self.entry_reposicao_descricao.get()
        self.ret_tamanho = self.entry_reposicao_tamanho.get()
        self.ret_quantidade = self.entry_reposicao_quantidade.get()

        self.msg_ret = messagebox.askyesno(
            '', 'DESEJA FAZER A CORREÇAO DO SEU ESTOQUE?')
        if self.msg_ret == TRUE:
            self.msg_ret1 = messagebox.askyesno(
                '', f'SERA RETIRADO DE SEU ESTOQUE  {self.ret_quantidade} UNIDADE '
                    f'\n PARA O SEGUIMENTO  {self.ret_seguimento}  DESEJA CONTINUAR?')
            if self.msg_ret1 == TRUE:
                self.adicionar_rep = self.cursor.execute(
                    f"""UPDATE  seguimento_venda  SET 
                                quantidade =quantidade - {self.ret_quantidade}
                                WHERE id = '{self.id}'
                                AND seguimento = '{self.ret_seguimento}'
                                AND tema = '{self.ret_tema}'
                                AND descricao = '{self.ret_descri}'
                                AND tamanho = '{self.ret_tamanho}'
                                                            """)
                self.cursor.commit()
                self.desconectar()
                messagebox.showinfo('', f'{self.ret_quantidade},{self.ret_seguimento} FOi RETIRADA ')
            else:
                messagebox.showinfo('', 'NENHUMA VALOR FOI RETIRADO FOI REALIZADA')
        else:
            messagebox.showinfo('', 'NENHUMA VALOR FOI RETIRADO FOI REALIZADA')
    #                            FUNÇAO EXPORTA VENDAS DOS MENUS                                                       #

    def expor_venda(self):
        import pandas as pd
        import pyodbc
        self.exporta_tudo = self.exp_entry_data.get()
        if self.exporta_tudo == '':
            messagebox.showinfo('', 'PREENCHA O CAMPO DATA PARA OBTER OS DADOS')
        else:
            self.exp_tudo = messagebox.askyesno(
                '', f'DESEJA EXPORTA TODOS OS SEGUIMENTOS REFERENTE A DATA {self.exporta_tudo}')
            if self.exp_tudo == TRUE:
                self.exp = messagebox.askyesno(
                    '', f'SERÁ EXPORTADOS TODOS OS SEGUIMENTOS REFERENTE A DATA {self.exporta_tudo} DESEJA CONTINUAR')
                if self.exp == TRUE:
                    self.dados_conexao = ("Driver={SQL Server};"
                                          "Server=DESKTOP-RFPLHG0;"
                                          "Database=seguimentos_cadastrados;")
                    messagebox.showinfo(
                        '', f'AS VENDAS REFERENTE A DATA {self.exporta_tudo} FOI EXPORTADA COM SUCESSO')

                    self.conexao = pyodbc.connect(self.dados_conexao)
                    self.cursor = self.conexao.cursor()

                    self.tabela = pd.read_sql(
                        f"SELECT * FROM seguimento_venda  WHERE data_venda='{self.exporta_tudo}'", self.conexao)

                    self.tabela1 = pd.DataFrame(data=self.tabela)

                    self.seguimento = self.tabela[
                        ['seguimento_venda', 'tema_venda', 'descricao_venda', 'tamanho_venda', 'quantidade_venda',
                         'valor_venda', 'total_venda']] \
                        .groupby(
                        ['seguimento_venda', 'tema_venda', 'descricao_venda', 'tamanho_venda', 'valor_venda']).sum()
                    self.seguimento.to_excel('PLANILHA DE VENDAS.xlsx')
                else:
                    messagebox.showinfo('', 'NENHUM VALOR SER EXPORTADO')
            else:
                messagebox.showinfo('', 'NENHUM VALOR SER EXPORTADO')

    def expor_selecao(self):
        import pandas as pd
        import pyodbc
        self.exp_data = self.exp_entry_data.get()
        self.selecao = self.entry_exp_venda_cliente.get()
        if self.exp_data == '' or self.selecao == '':
            messagebox.showinfo('', 'PREENCHA O CAMPO DATA E CLIENTE PARA OBTER OS DADOS')
        else:
            self.meg_selecao = messagebox.askyesno(
                '', f'DESEJA EXPORTAR  AS VENDAS  DA DATA {self.exp_data} REFETENTE A {self.selecao}')
            if self.meg_selecao == TRUE:
                self.exp = messagebox.askyesno(
                    '', f'SERÁ EXPORTADOS TODAS AS VENDAS REFERENTE A DATA {self.exp_data} '
                        f'DO CLIENTE {self.selecao} DESEJA CONTINUAR')
                if self.exp == TRUE:
                    self.dados_conexao = (
                        "Driver={SQL Server};"
                        "Server=DESKTOP-RFPLHG0;"
                        "Database=seguimentos_cadastrados;")
                    messagebox.showinfo(
                        '', f'AS VENDAS REFERENTE A DATA {self.exp_data}  DO CLIENTE {self.selecao}'
                            f'\nFOI EXPORTADA COM SUCESSO')

                    self.conexao = pyodbc.connect(self.dados_conexao)
                    self.cursor = self.conexao.cursor()

                    self.planilha = pd.read_sql(
                        f"SELECT * FROM seguimento_venda   WHERE data_venda='{self.exp_data}' "
                        f"          and cliente = '{self.selecao}'", self.conexao)

                    self.planilha1 = pd.DataFrame(data=self.planilha)

                    self.seguimento_exp = self.planilha[['data_venda', 'cliente', 'seguimento_venda',
                                                         'tema_venda',
                                                         'descricao_venda',
                                                         'tamanho_venda',
                                                         'quantidade_venda',
                                                         'valor_venda', ''
                                                                        'total_venda']] \
                        .groupby(['data_venda', 'cliente', 'seguimento_venda',
                                  'tema_venda',
                                  'descricao_venda',
                                  'tamanho_venda', 'valor_venda']).sum()
                    self.seguimento_exp.to_excel('PLANILHA_SELEÇAO.xlsx')
                else:
                    messagebox.showinfo('', 'NENHUM VALOR SER EXPORTADO')
            else:
                messagebox.showinfo('', 'NENHUM VALOR SER EXPORTADO')

    def exporta_agrupamento(self):
        import pandas as pd
        import pyodbc
        self.exp_data_exel = self.exp_entry_data.get()
        if self.exp_data_exel == '':
            messagebox.showinfo('', 'PREENCHA O CAMPO DATA E CLIENTE PARA OBTER OS DADOS')
        else:
            self.mens_selecao = messagebox.askyesno('',
                                                    f'DESEJA EXPORTAR TODAS  AS VENDAS  DA DATA {self.exp_data_exel}')
            if self.mens_selecao == TRUE:
                self.exporta = messagebox.askyesno(
                    '', f'SERÁ EXPORTADA TODAS AS VENDAS REFERENTE A DATA {self.exp_data_exel}\n'
                        f'E TODOS OS VALORES NELES OBTIDOS SERAO AGRUPADOS EM CADA SEGUIEMNTO\n'
                        f'DESEJA CONTINUAR?')
                if self.exporta == TRUE:
                    self.dados_conexao = ("Driver={SQL Server};"
                                          "Server=DESKTOP-RFPLHG0;"
                                          "Database=seguimentos_cadastrados;")
                    messagebox.showinfo('', f'TODOS OS DADOS FORAM AGRUPADO COM SUCESSO {self.exp_data_exel}')

                    self.conexao = pyodbc.connect(self.dados_conexao)
                    self.cursor = self.conexao.cursor()

                    self.planilhaa = pd.read_sql(
                        f"SELECT * FROM seguimento_venda   WHERE data_venda='{self.exp_data_exel}' ", self.conexao)

                    self.planilhaa1 = pd.DataFrame(data=self.planilhaa)

                    self.seguimento_expp = self.planilhaa[
                        ['data_venda', 'seguimento_venda', 'tamanho_venda', 'quantidade_venda', ]] \
                        .groupby(['data_venda', 'seguimento_venda', 'tamanho_venda']).sum()
                    self.seguimento_expp.to_excel('PLANILHA_AGRUPADA.xlsx')
                else:
                    messagebox.showinfo('', 'NENHUM VALOR SER EXPORTADO')
            else:
                messagebox.showinfo('', 'NENHUM VALOR SER EXPORTADO')

    #                                    FUNCAO DE PLANILHA                                                            #
    # TELA PRINCIPAL

    def planila(self):
        self.app11 = Toplevel()
        self.app11.geometry('1350x600+0+56')
        self.app11.resizable(False, False)
        self.app11.transient(self.app)
        self.app11.focus_force()
        self.app11.grab_set()
        self.app11.configure(bg='#2a2a29')

        #                                    FRAMES SUPERIOR E INFERIOR                                                #

        self.frame_planilha = Frame(self.app11, bg="white", width=885, height=290, borderwidth=9, relief="raised",
                                    background='#7f7f7f')
        self.frame_planilha.place(x=10, y=10)

        self.frame_planilha_inferior = Frame(self.app11, bg="white", width=885, height=260, borderwidth=9,
                                             relief="raised",
                                             background='#7f7f7f')
        self.frame_planilha_inferior.place(x=10, y=310)

        self.lb_planilha_cadastro = Label(
            self.app11, text='ARÉA DE PEDIDOS', bg='#574c26',
            fg='white', font='times 15', relief='raised')
        self.lb_planilha_cadastro.place(x=300, y=23)

        #                                  LABEL  EXPORTA  PLANILHA                                                #

        self.frame_exporta_cadastro = Frame(
            self.app11, bg="white", width=450, height=559, borderwidth=9,
            relief="raised",
            background='#7f7f7f')
        self.frame_exporta_cadastro.place(x=900, y=10)
        self.lb_exporta_cadastro = Label(self.app11, text='EXPORTAR PLANILHA', bg='#574c26',
                                 fg='white', font='times 15', relief='raised')
        self.lb_exporta_cadastro.place(x=1050, y=23)

        #                          ENTRYS   EXPORTA PLANILHA     LADO DIREITO                                 #
        self.clientee = Entry(self.app11, font='times 15 bold ', width=15, relief=FLAT)
        self.clientee.place(x=920, y=80)
        self.seguimento = Entry(self.app11, font='times 15 bold ', width=15, relief=FLAT)
        self.seguimento.place(x=920, y=120)

        #                              LABEL EXPORTA PLANILHA   LADO DIREITO                                           #

        self.label_cliente = Label(self.app11,text='CLIENTE',
                                   font='times 15 bold',
                                   relief='raised',bg='#7f7f7f',fg='#ffffff')
        self.label_cliente.place(x=1040, y=80)

        self.label_seguimento = Label(self.app11,text='SEGUIMENTO',
                                      font='times 15 bold',
                                      relief='raised',
                                      bg='#7f7f7f',fg='#ffffff')
        self.label_seguimento.place(x=1040, y=120)

#                                     BOTOES  EXPORTA PLANILHA  LADO DIREITO                                                #
        self.imag3 = PhotoImage(file='imagens/corte.png')
        self.botao_data = Button(self.app11, text='EXPORTA',compound=TOP,image=self.imag3,bg='#7f7f7f',font='times',
                                 relief=FLAT,fg='orange',
                                 command=self.exporta_planilha_corte)
        self.botao_data.place(x=1140, y=165)

        self.img= PhotoImage(file='imagens/cliente.png')
        self.botao_cliente = Button(self.app11, text='EXPORTA',image=self.img,compound=TOP
                                    ,fg='orange',font='times',bg='#7f7f7f',relief=FLAT,
                                    command=self.exporta_planilha_cliente)
        self.botao_cliente.place(x=920, y=180)

        self.img1= PhotoImage(file='imagens/seguimento.png')
        self.botao_seguimento = Button(self.app11, text='EXPORTA',image=self.img1,
                                       compound=TOP,font='times',bg='#7f7f7f',relief=FLAT,fg='orange',
                                       command=self.exporta_planilha_seguimento)
        self.botao_seguimento.place(x=1030, y=165)

        self.planilhaP= PhotoImage(file='imagens/exel.png')
        self.botao_planilha = Button(self.app11, text='PEDIDOS',image=self.planilhaP ,compound=TOP,
                                        font='times', bg='#7f7f7f', relief=FLAT, fg='orange',
                                       command=self.exporta_planilha_pedido)
        self.botao_planilha.place(x=1250, y=165)

#                                                      ENTRY CALADAERIO  LADO DIREITO                                 #
        self.data_exp = DateEntry(self.app11, dateformat=3, width=12, background='darkblue',
                                    foreground='white', borderwidth=4, locale='pt_BR',
                                    date_pattern='d/m/y', font='TIMES 12 bold')  # en_US Calendar =2019)
        self.data_exp.place(x=920, y=40)

        self.valorPedido=PhotoImage(file='imagens/total.png')

        self.valor_pedido=Button(self.app11,text= 'VALOR TOTAL',image=self.valorPedido,
                                 font='times',bg='#7f7f7f',relief=FLAT,fg='orange',
                                 command=self.valorPedidoPlanilha)
        self.valor_pedido.place(x=790,y=210)



        #          AQUI SAO AS LABEL ORDENADA IGUAL AS COLUNAS DO BANCO DE DADOS                                       #
        self.lb_planilha_cliente = Label(self.app11, text='CLIENTE', bg='#7f7f7f',
                                    fg='white', relief='raised', font='times 14 bold')
        self.lb_planilha_cliente.place(x=180, y=25)

        self.lb_planilha_id = Label(self.app11, text='ID', bg='#7f7f7f',
                                    fg='white', relief=FLAT, font='times 14 bold')
        self.lb_planilha_id.place(x=475, y=60)

        self.lb_planilha_seguimento = Label(self.app11, text='SEGUIMENTO', bg='#7f7f7f',
                                            fg='white', font='times 14', relief='raised')
        self.lb_planilha_seguimento.place(x=290, y=60)

        self.lb_planilha_tema = Label(self.app11, text='TEMA/COR(SEG)', font='times 14', bg='#7f7f7f',
                                      fg='white', relief='raised')
        self.lb_planilha_tema.place(x=290, y=100)

        self.lb_planilha_descricao = Label(self.app11, text='DESCRIÇAO', bg='#7f7f7f',
                                           fg='white', relief='raised', font='times 14')
        self.lb_planilha_descricao.place(x=290, y=140)

        self.lb_planilha_tamanho = Label(self.app11, text='TAMANHO', bg='#7f7f7f',
                                         fg='white', relief='raised', font='times 14')
        self.lb_planilha_tamanho.place(x=120, y=175)

        self.lb_planilha_quantidade = Label(self.app11, text='QUANTIDADE', bg='#7f7f7f',
                                            fg='white', relief='raised', font='times 14')
        self.lb_planilha_quantidade.place(x=120, y=212)

        self.lb_planilha_valor = Label(self.app11, text='VALOR', bg='#7f7f7f',
                                       fg='white', relief='raised', font='times 14')
        self.lb_planilha_valor.place(x=120, y=252)
        self.treeview_planilha1()

        #                   AQUI SAO MEUS ENTRYS TODOS ELES SEGUEM O MSM PADRAO DO BANCO DE DADOS                    #

        self.planilha_id = Entry(self.app11, font='times 15 bold ', width=5,    background='#7f7f7f', fg='white',
                                       relief=FLAT)
        self.planilha_id.place(x=420, y=60)

        self.data_planilha = DateEntry(self.app11, dateformat=3, width=12, background='darkblue',
                                    foreground='white', borderwidth=4, locale='pt_BR',
                                    date_pattern='d/m/y')
        self.data_planilha.place(x=500, y=30)

        self.planilha_cliente = Entry(self.app11, font='times 15', width=15)
        self.planilha_cliente.place(x=25, y=25)
        self.planilha_seguimento = Entry(self.app11, font='times 15', width=25)
        self.planilha_seguimento.place(x=30, y=60)
        self.planilha_tema = Entry(self.app11, font='times 15', width=25)
        self.planilha_tema.place(x=30, y=100)
        self.planilha_descricao = Entry(self.app11, font='times 15', width=25)
        self.planilha_descricao.place(x=30, y=140)
        self.planilha_tamanho = Entry(self.app11, font='times 15', width=7)
        self.planilha_tamanho.place(x=30, y=175)
        self.planilha_quantidade = Entry(self.app11, font='times 15 ', width=7)
        self.planilha_quantidade.place(x=30, y=213)
        self.planilha_valor = Entry(self.app11, font='times 15 ', width=7)
        self.planilha_valor.place(x=30, y=253)

        # LABEL ESQUEDAR DA FUNCAO AREA DE PEDIDOS                                                #

        self.lab_planilha = Label(self.app11, text='BUSCA CLIENTE', bg='#6b6a68', fg='white', relief='raised',
                                  font='times 15')
        self.lab_planilha.place(x=600, y=70)

        self.lab_planilha_tema = Label(self.app11, text='SEGUIMENTO', bg='#6b6a68', fg='white', relief='raised',
                                       font='times 15')
        self.lab_planilha_tema.place(x=690, y=135)

        self.cliente = Label(self.app11, text='PESQUISA POR CLIENTE', bg='#6b6a68', fg='white', relief='raised',
                                       font='times 15')
        self.cliente.place(x=1030, y=250)

        self.nome = Label(self.app11, text='NOME', bg='#6b6a68', fg='white', relief='raised',
                             font='times 15')
        self.nome.place(x=1070, y=285)

        # ENTRY DA PESQUISA DE CLIENTE DA FUNÇAO PLANILHA LADO DIREITO

        self.nomecliente= Entry(self.app11,font='times 15 ', width=15)
        self.nomecliente.place(x=910, y=285)

        self.procura= Button(self.app11,text='BUSCA',bg='orange',command=self.buscaclienteplanilha)
        self.procura.place(x=1140, y=285)
        self.treview_planilha_exporta()

        self.entry_planilha_pesquisa = Entry(self.app11, font='times 15', width=25)
        self.entry_planilha_pesquisa.place(x=510, y=105)

        self.entry_planilha_pes_seguimento = Entry(self.app11, font='times 15', width=12)
        self.entry_planilha_pes_seguimento.place(x=510, y=135)
        #                               BOTAO QUE REALIZA A BUSCA PARA ALTERAR VALOR                        #
        self.planilha_vendas = Button(self.app11, text='BUSCAR', bg='#574c26', fg='white', font='times 12',
                                      command=self.pesquisar_item_planilha)
        self.planilha_vendas.place(x=710, y=170)

        self.bt_planilha_pesquisar = Button(self.app11, text='CADASTRAR', bg='#f6f3ea', pady=5,
                                            font='times 13 bold',
                                            command=self.inserir_dados_planilha)
        self.bt_planilha_pesquisar.place(x=300, y=200)

        self.bt_planilha_corrigir = Button(self.app11, text='CORRIGIR', bg='#574c26', fg='white', pady=5,
                                           font='times 13 bold',
                                           command=self.alterar_valores_planilha)
        self.bt_planilha_corrigir.place(x=700, y=250)

        self.bt_planilha_alterar = Button(self.app11, text='ALTERAR VALOR', bg='red', fg='white', pady=5,
                                           font='times 13 bold', command=self.alterar_valor)
        self.bt_planilha_alterar.place(x=420, y=250)

        self.bt_planilha_voltar = Button(self.app11, text='VOLTAR', command=self.app11.destroy, bg='#7f7f7f',
                                         fg='white',
                                         relief=FLAT, pady=5,
                                         font='times 15')
        self.bt_planilha_voltar.place(x=680, y=20)

    def treeview_planilha1(self):
        self.columns = (
            'id', 'data_pedido', 'cliente', 'seguimento', 'tema', 'descricao', 'tamanho', 'quantidade', 'valor')
        self.treeview_planilha = ttk.Treeview(self.app11, columns=self.columns, show='headings')
        self.treeview_planilha.heading('id', text='ID')
        self.treeview_planilha.heading('data_pedido', text='Data')
        self.treeview_planilha.heading('cliente', text= 'Cliente')
        self.treeview_planilha.heading('seguimento', text='Seguimento')
        self.treeview_planilha.heading('tema', text='Tema')
        self.treeview_planilha.heading('tamanho', text='Tamanho')
        self.treeview_planilha.heading('descricao', text='Descriçao')
        self.treeview_planilha.heading('quantidade', text='QTDD')
        self.treeview_planilha.heading('valor', text='Valor')

        self.treeview_planilha.column('id', width=30)
        self.treeview_planilha.column('data_pedido', width=90)
        self.treeview_planilha.column('cliente', width=80)
        self.treeview_planilha.column('seguimento', width=200)
        self.treeview_planilha.column('tema', width=150)
        self.treeview_planilha.column('tamanho', width=50)
        self.treeview_planilha.column('descricao', width=150)
        self.treeview_planilha.column('quantidade', width=40)
        self.treeview_planilha.column('valor', width=50)
        self.treeview_planilha.place(x=20, y=320, width=840, height=240)
        self.scroll_planilha = Scrollbar(self.app11, orient=VERTICAL, command=self.treeview_planilha.yview)
        self.treeview_planilha.configure(yscrollcommand=self.scroll_planilha.set)
        self.scroll_planilha.place(x=865, y=320, width=25, height=240)
        self.treeview_planilha.bind("<Double-1>", self.double_click_planilha)

    def inserir_dados_planilha(self):
        self.sql()
        self.treeview_planilha.delete(*self.treeview_planilha.get_children())
        self.data = self.data_planilha.get()
        self.plan_cliente = self.planilha_cliente.get()
        self.plan_segui = self.planilha_seguimento.get().capitalize()
        self.plan_tema = self.planilha_tema.get().capitalize()
        self.plan_descricao = self.planilha_descricao.get().capitalize()
        self.plan_tamanho = self.planilha_tamanho.get().capitalize()
        self.plan_quantidade = self.planilha_quantidade.get()
        self.plan_valor = self.planilha_valor.get().replace(',', '.')
        try:
            if self.plan_segui == '' or \
                    self.plan_tema == '' or \
                    self.plan_descricao == '' or \
                    self.plan_tamanho == '' or \
                    self.plan_quantidade == '' or \
                    self.plan_valor == '':
                messagebox.showinfo('PREENCHA!', 'PREECHA OS CAMPOS CORRETAMENTE PARA REALIZAR O CADASTRO')
            else:
                self.mensg = messagebox.askyesno('CADASTRO', f'DESEJA  CADASTRAR O SEGUIMENTO {self.plan_segui}?')

                if self.mensg == TRUE:
                    messagebox.showinfo('CONCLUIDO', 'SEGUIMENTO CADASTRADO COM SUCESSO')
                    self.planilha = self.cursor.execute(
                            f"""INSERT INTO planilha 
                                (data_pedido,cliente,seguimento,tema,descricao,tamanho,quantidade,valor)
                                                      VALUES
                                                      ('{self.data}',
                                                      '{self.plan_cliente}',
                                                      '{self.plan_segui}',
                                                      '{self.plan_tema}',
                                                      '{self.plan_descricao}',
                                                      '{self.plan_tamanho}',
                                                      '{self.plan_quantidade}',
                                                      {self.plan_valor})
                                                      """)

                    self.p = self.cursor.execute(
                        f""" SELECT data_pedido,id,cliente,seguimento,tema,descricao,tamanho,quantidade,valor
                            FROM planilha
                            WHERE cliente= '{self.plan_cliente}'
                            and data_pedido = '{self.data}' """)

                    for linha in self.cursor.fetchall():
                        self.v =\
                            linha.id,\
                            linha.data_pedido,\
                            linha.cliente,\
                            linha.seguimento,\
                            linha.tema,\
                            linha.descricao,\
                            linha.tamanho,\
                            linha.quantidade, \
                            linha.valor
                        self.treeview_planilha.insert('', 'end', values=self.v)
                    self.cursor.commit()
                    self.desconectar()

                    self.mensg1 = messagebox.askyesno(
                        '', f'DESEJA CONTINUAR A CADASTRAR O SEGUIMENTO {self.plan_segui}')
                    if self.mensg1 == TRUE:
                        self.planilha_tema.delete(0, END)
                        self.planilha_descricao.delete(0, END)
                        self.planilha_tamanho.delete(0, END)

                    else:
                        messagebox.showinfo(
                            '', 'OKAY, PARA REALIZAR UM NOVO CADASTRO E NECESSARIO PREENCHER TODOS OS CAMPOS NOVAMENTE ')

                else:
                    messagebox.showinfo('', 'NENHUM SEGUIMENTO SERÁ CADASTRADO')
        except:
            pass

    def pesquisar_item_planilha(self):
        self.treeview_planilha.delete(*self.treeview_planilha.get_children())
        self.sql()
        self.cliente = self.entry_planilha_pesquisa.get()
        self.busca_seguimento = self.entry_planilha_pes_seguimento.get()
        self.pes_data = self.data_planilha.get()

        if self.busca == '':
            messagebox.showinfo('PESQUISA DE SEGUIMENTO', 'DIGITE O SEGUIMENTO QUE DESEJA REALIZAR A BÚSCA')
        elif self.cliente != '' and self.busca_seguimento != '' and self.pes_data != '':
            self.pesquisar1 = self.cursor.execute(
                f"""SELECT id,data_pedido,cliente,seguimento,tema,descricao,tamanho,quantidade,valor
                    FROM planilha
                    WHERE  cliente like '{self.cliente}%'
                    AND data_pedido = '{self.pes_data}'
                    AND seguimento ='{self.busca_seguimento}'""")

            for linha in self.pesquisar1:
                self.buscar1 = linha.id,\
                               linha.data_pedido,\
                               linha.cliente,\
                               linha.seguimento,\
                               linha.tema,\
                               linha.descricao,\
                               linha.tamanho, \
                               linha.quantidade,\
                               linha.valor
                self.treeview_planilha.insert('', 'end', values=self.buscar1)

        elif self.cliente != '' and self.pes_data != '':
            self.pesquisar1 = self.cursor.execute(
                f"""SELECT id,data_pedido,cliente,seguimento,tema,descricao,tamanho,quantidade,valor
                    FROM planilha
                    WHERE  cliente like '{self.cliente}%'
                    AND data_pedido = '{self.pes_data}' """)

            for linha in self.pesquisar1:
                self.buscar1 = linha.id,\
                               linha.data_pedido,\
                               linha.cliente,\
                               linha.seguimento,\
                               linha.tema,\
                               linha.descricao,\
                               linha.tamanho, \
                               linha.quantidade,\
                               linha.valor
                self.treeview_planilha.insert('', 'end', values=self.buscar1)
        else:
            messagebox.showinfo('','PARA REALIZAR UMA BUSCA PRENCHA CORRETAMENTE OS CAMPO')

    def alterar_valores_planilha(self):
        self.sql()
        # self.treeview_planilha.delete(*self.treeview_planilha.get_children())
        self.plani_id = self.planilha_id.get()
        self.plani_cliente = self.planilha_cliente.get()
        self.plani_tamanho = self.planilha_tamanho.get().capitalize()
        self.plani_tema = self.planilha_tema.get().capitalize()
        self.plani_descricao = self.planilha_descricao.get().capitalize()
        self.plani_valor = self.planilha_valor.get()
        self.plani_seguimento = self.planilha_seguimento.get().capitalize()
        self.plani_quantidade = self.planilha_quantidade.get()

        if self.planilha_cliente != '' and \
                self.planilha_tema != '' and \
                self.planilha_descricao != '' and \
                self.planilha_tamanho != '' and self.planilha_quantidade != '':
            self.corrigir1 = self.cursor.execute(
                f"""UPDATE planilha SET cliente = '{self.plani_cliente}',
                 seguimento  = '{self.plani_seguimento}',
                 tema = '{self.plani_tema}',
                 descricao = '{self.plani_descricao}',
                 tamanho = '{self.plani_tamanho}',
                 quantidade = '{self.plani_quantidade}' 
                                                        WHERE id = '{self.plani_id}'
                                                                 """)
            messagebox.showinfo('ALTERAÇAO','DADOS ALTERADO COM SUCESSO')
        else:
            messagebox.showwarning('AVISO !','IMPORTANTE INFORMAR OS VALORES CORRETAMENTE PARA REALIZAR  ALTERAÇAO')


        self.cursor.commit()
        self.desconectar()

    def alterar_valor(self):
        self.sql()
        # self.treeview_planilha.delete(*self.treeview_planilha.get_children())
        self.planil_id = self.planilha_id.get().strip()
        self.planil_cliente = self.planilha_cliente.get().strip()
        self.planil_tamanho = self.planilha_tamanho.get().strip().capitalize()
        self.planil_tema = self.planilha_tema.get().strip().capitalize()
        self.planil_descricao = self.planilha_descricao.get().capitalize()
        self.planil_valor = self.planilha_valor.get().strip().replace(',', '.')
        self.planil_seguimento = self.planilha_seguimento.get().strip().capitalize()
        self.planil_quantidade = self.planilha_quantidade.get().strip().replace(',', '.')

        if self.planil_valor != '' \
                or self.planil_cliente != ''\
                or self.planil_id != ''\
                or self.planil_tema != ''\
                or self.planil_descricao != '' \
                or self.planil_seguimento != '' :
            self.novopreco= messagebox.askyesno('', f'DESEJA ALTERAR O VALOR DO SEGUIMENTO {self.planil_seguimento}')
            try:
                if self.novopreco == TRUE:
                    self.preco = self.cursor.execute(
                            f"""UPDATE planilha SET valor = '{self.planil_valor}'
                                            WHERE cliente = '{self.planil_cliente}'
                                            AND seguimento = '{self.planil_seguimento}'
                                            AND tema = '{self.planil_tema}'
                                            AND tamanho= '{self.planil_tamanho}'
                                            AND descricao = '{self.planil_descricao}'
                                            AND id = '{self.planil_id}'""")
                    messagebox.showinfo('', 'VALOR ALTERADO COM SUCESSO')
                else:
                    messagebox.showinfo('', 'NENHUM VALOR FOI ALTERADO')
            except:
                messagebox.showerror(
                    'ERRO !',f' O VALOR {self.planil_valor} ESTÁ INVALIDO POR FAVOR INFORME O VALOR CORRETAMENTE')
        else:
            messagebox.showwarning(
                'AVISO !',f'INFORME O NOVO PREÇO QUE SERÁ DEFINIDO PARA O SEGUIMENTO {self.planil_seguimento} ')
        self.cursor.commit()
        self.desconectar()

    def valorPedidoPlanilha(self):
        self.pla_cliente = self.planilha_cliente.get()
        self.pla_data= self.data_planilha.get()
        try:
            if self.pla_cliente and self.pla_data != '':
                somar=self.cursor.execute(f"""SELECT SUM(total) FROM planilha
                WHERE data_pedido = '{self.pla_data}'
                AND cliente = '{self.pla_cliente}'
                """)
                for total in somar:
                    valor=total[0]
                    messagebox.showinfo('',f'O VALOR TOTAL A PAGAR {valor:,.2f}')
            else:
                messagebox.showinfo('','INFORME O NOME DO CLIENTE EA DATA CORRETAMENTE')
        except:
            messagebox.showwarning('AVISO',f'NENHUM VALOR ENCONTARDO PARA {self.pla_cliente},\n'
                                           f' REFERENTE A DATA {self.pla_data}')

    def limpa_campo_planilha(self):
        self.data_planilha.delete(0, END)
        self.planilha_id.delete(0, END)
        self.planilha_cliente.delete(0, END)
        self.planilha_seguimento.delete(0, END)
        self.planilha_tema.delete(0, END)
        self.planilha_descricao.delete(0, END)
        self.planilha_tamanho.delete(0, END)
        self.planilha_quantidade.delete(0, END)
        self.planilha_valor.delete(0, END)

    def double_click_planilha(self, event):
        self.limpa_campo_planilha()
        self.treeview_planilha.selection()
        for y in self.treeview_planilha.selection():
            col1, col2, col3, col4, col5, col6, col7, col8, col9 = self.treeview_planilha.item(y, 'values')
            self.planilha_id.insert(END, col1)
            self.data_planilha.insert(END, col2)
            self.planilha_cliente.insert(END, col3)
            self.planilha_seguimento.insert(END, col4)
            self.planilha_tema.insert(END, col5)
            self.planilha_descricao.insert(END, col6)
            self.planilha_tamanho.insert(END, col7)
            self.planilha_quantidade.insert(END, col8)
            self.planilha_valor.insert(END, col9)
            self.planilha_valor.delete(0, END)

#  FUNÇAO DE EXPORTA TODAS AS VENDAS DO DIA  EM UMA UNICA PLANILHA

    def exporta_planilha_pedido(self):
        import pandas as pd
        import pyodbc
        self.exp_data_planilha = self.data_exp.get()
        try:
            if self.exp_data_planilha == '':
                messagebox.showinfo('', 'PREENCHA O CAMPO DATA E  PARA OBTER OS DADOS')
            else:

                self.mensa_selecao = messagebox.askyesno(
                    '', f'DESEJA EXPORTAR TODAS  AS VENDAS  DA DATA {self.exp_data_planilha}')
                if self.mensa_selecao == TRUE:
                    self.exporta = messagebox.askyesno(
                        '', f'SERÁ EXPORTADA TODAS AS VENDAS REFERENTE A DATA {self.exp_data_planilha}\n'
                            f'E TODOS OS VALORES NELES OBTIDOS SERAO AGRUPADOS EM CADA SEGUIEMENTO\n'
                            f'DESEJA CONTINUAR?')
                    if self.exporta == TRUE:
                        self.dados_conexao = ("Driver={SQL Server};"
                                              "Server=DESKTOP-RFPLHG0;"
                                              "Database=seguimentos_cadastrados;")

                        self.conexao = pyodbc.connect(self.dados_conexao)
                        self.cursor = self.conexao.cursor()

                        self.planilhaa = pd.read_sql(
                            f"SELECT * FROM planilha WHERE data_pedido='{self.exp_data_planilha}'", self.conexao)

                        self.planilhaa1 = pd.DataFrame(data=self.planilhaa)

                        self.seguimento_expp = self.planilhaa[
                            ['data_pedido', 'seguimento', 'tema','descricao', 'tamanho', 'quantidade','valor' ]] \
                            .groupby(['data_pedido', 'seguimento', 'descricao','tamanho']).sum()
                        self.seguimento_expp.to_excel('PLANILHA_DE_PEDIDO.xlsx')
                        messagebox.showinfo('', f'TODOS OS DADOS FORAM AGRUPADO COM SUCESSO {self.exp_data_planilha}')
                    else:
                        messagebox.showinfo('CANCELADO!', 'NENHUMA EXPORTAÇAO FOI REALIZADA')
                else:
                    messagebox.showinfo('CANCELADO!', 'NENHUMA EXPORTAÇAO FOI REALIZADA')
        except:
            messagebox.showerror('ERRO!', f'SERÁ EXPORTADO UMA TABELA VAZIA POIS NA HÁ VALORES'
                                          f' REFERENTE A DATA {self.exp_data_planilha} \n'
                                          f'VERIFIQUE SE INFORMOU OS DADOS CORRETAMENTE')

# FUNÇAO PARA EXPORTA O SEGUIMENTO E QUANTIDADE PARA REALIZAR UM CORTE

    def exporta_planilha_corte(self):
        import pandas as pd
        import pyodbc
        self.expo_data_planilha = self.data_exp.get()
        try:
            if self.expo_data_planilha == '':
                messagebox.showinfo('', 'PREENCHA O CAMPO DATA E  PARA OBTER OS DADOS')
            else:

                self.mensa_selecao = messagebox.askyesno(
                    '', f'DESEJA EXPORTAR TODOS  OS SEGUIMENTO PARA CORTE REFERENTE A DATA  {self.expo_data_planilha}')
                if self.mensa_selecao == TRUE:
                    self.exporta = messagebox.askyesno(
                        '', f'SERÁ EXPORTADA TODOS OS SEGUIMENTOS REFERENTE A DATA {self.expo_data_planilha}\n'
                            f'E TODOS OS VALORES NELES OBTIDOS SERAO AGRUPADOS EM CADA SEGUIEMENTO\n'
                            f'DESEJA CONTINUAR?')
                    if self.exporta == TRUE:
                        self.dados_conexao = ("Driver={SQL Server};"
                                              "Server=DESKTOP-RFPLHG0;"
                                              "Database=seguimentos_cadastrados;")

                        self.conexao = pyodbc.connect(self.dados_conexao)
                        self.cursor = self.conexao.cursor()

                        self.planilhaa = pd.read_sql(
                            f"SELECT * FROM planilha WHERE data_pedido='{self.expo_data_planilha}'", self.conexao)

                        self.planilhaa1 = pd.DataFrame(data=self.planilhaa)

                        self.seguimento_expp = self.planilhaa[
                            ['data_pedido', 'seguimento','tema', 'tamanho', 'quantidade', ]] \
                            .groupby(['data_pedido', 'seguimento', 'tamanho']).sum()
                        self.seguimento_expp.to_excel('PLANILHA_DE_CORTE.xlsx')
                        messagebox.showinfo('', f'TODOS OS DADOS FORAM AGRUPADO COM SUCESSO {self.expo_data_planilha}')
                    else:
                        messagebox.showinfo('CANCELADO!', 'NENHUMA EXPORTAÇAO FOI REALIZADA')
                else:
                    messagebox.showinfo('CANCELADO!', 'NENHUMA EXPORTAÇAO FOI REALIZADA')
        except:
            messagebox.showerror('ERRO!', f'SERÁ EXPORTADO UMA TABELA VAZIA POIS NA HÁ VALORES'
                                          f' REFERENTE A DATA {self.expo_data_planilha} \n'
                                          f'VERIFIQUE SE INFORMOU OS DADOS CORRETAMENTE')

# FUNÇAO DE EXPORTA VENDA DO CLIENTE INDIVIDUALMENTE

    def exporta_planilha_cliente(self):
        import pandas as pd
        import pyodbc
        self.cliente_planilha = self.clientee.get()
        self.dataa = self.data_exp.get()

        try:
            if self.cliente_planilha != '' and self.dataa != '':
                self.mens_exporta_cliente = messagebox.askyesno(
                    '', f'DESEJA EXPORTA A VENDA DO DIA  {self.dataa} DO CLIENTE  {self.cliente_planilha} ?')
                if self.mens_exporta_cliente == TRUE:
                    self.dados_conexao = ("Driver={SQL Server};"
                                          "Server=DESKTOP-RFPLHG0;"
                                          "Database=seguimentos_cadastrados;")
                    self.conexao = pyodbc.connect(self.dados_conexao)
                    self.cursor = self.conexao.cursor()
                    self.planilhaa = pd.read_sql(
                        f" SELECT * FROM planilha WHERE data_pedido='{self.dataa}'"
                        f"                          AND cliente = '{self.cliente_planilha}'", self.conexao)

                    self.planilhaa1 = pd.DataFrame(data=self.planilhaa)

                    self.seguimento_expp = self.planilhaa[
                        ['data_pedido', 'cliente', 'seguimento', 'tema','descricao','tamanho', 'quantidade', 'valor', 'total' ]] \
                        .groupby(['data_pedido', 'cliente', 'seguimento', 'tema', 'descricao', 'tamanho']).sum()
                    self.seguimento_expp.to_excel('PLANILHA_INDIVIDUAL.xlsx')
                    messagebox.showinfo('CONCLUIDO','EXPORTAÇAO BEM SUCEDIDA')
                else:
                    messagebox.showinfo('CANCELADO', 'NENHUMA EXPORTAÇAO FOI REALIZADA')
            else:
                messagebox.showinfo(
                    '', 'INFORME O NOME DO CLIENTE QUE DESEJA EXPORTAR')
        except:
            messagebox.showerror('ERRO!',f'SERÁ EXPORTADO UMA TABELA VAZIA POIS NA HÁ VALORES PARA {self.cliente_planilha}\n'
                                         f' REFERENTE A DATA {self.dataa} \n'
                                         f'VERIFIQUE SE INFORMOU OS DADOS CORRETAMENTE')

# FUNÇAO DE EXPORTA VENDAS POR SEGUIMENTOS DE TODOS OS PEDIDOS DO DIA

    def exporta_planilha_seguimento(self):
        import pandas as pd
        import pyodbc
        self.p_seguimento = self.seguimento.get()
        self.exp_da = self.data_exp.get()
        try:
            if self.exp_da != '' and self.p_seguimento != '':
                self.mens_expor_seguimento = messagebox.askyesno(
                    '', f'DESEJA EXPORTA TODOS OS SEGUIMENTO REFENTE A DATA {self.exp_da}\n'
                        f' REFERENTE AO SEGUIMENTO {self.p_seguimento} ?')
                if self.mens_expor_seguimento == TRUE:
                    self.mens_expor_seg = messagebox.askyesno(
                        '', f'SERÁ EXPORTADOS TODOS OS SEGUIMENTO REALIZADO NA DATA {self.exp_da} \n'
                            f'DE TODOS OS CLIENTE DESEJA COPNTINUAR ?')
                    if self.mens_expor_seg == TRUE:
                        self.dados_conexao = ("Driver={SQL Server};"
                                              "Server=DESKTOP-RFPLHG0;"
                                              "Database=seguimentos_cadastrados;")
                        self.conexao = pyodbc.connect(self.dados_conexao)
                        self.cursor = self.conexao.cursor()
                        self.planilhaa = pd.read_sql(
                            f"SELECT * FROM planilha  WHERE data_pedido = '{self.exp_da}'"
                            f"AND seguimento = '{self.p_seguimento}'", self.conexao)
                        self.planilha1 = pd.DataFrame(data=self.planilhaa)
                        self.seguimento_expp = self.planilhaa[
                            ['data_pedido', 'seguimento','tema','descricao', 'tamanho', 'quantidade','valor','total' ]]\
                            .groupby(['data_pedido', 'seguimento','tema','descricao', 'tamanho']).sum()
                        self.seguimento_expp.to_excel('PLANILHA_seguimento.xlsx')

                        messagebox.showinfo('','EXPORTAÇAO BEM SUCEDIDA')
                    else:
                        messagebox.showinfo('CANCELADO', 'NENHUMA EXPORTAÇAO FOI REALIZADA')
                else:
                    messagebox.showinfo('CANCELADO', 'NENHUMA EXPORTAÇAO FOI REALIZADA')

            else:
                messagebox.showinfo('', 'INFORME O NOME DO SEGUIMENTO QUE DESEJA EXPORTA')
        except:
            messagebox.showerror('ERRO!', f'SERÁ EXPORTADO UMA TABELA VAZIA POIS NA HÁ VALORES'
                                          f' REFERENTE A DATA {self.exp_da} \n'
                                          f'PARA  SEGUIMENTO {self.p_seguimento} '
                                          f'VERIFIQUE SE INFORMOU OS DADOS CORRETAMENTE')

# treview do lado direto da tela planilha

    def treview_planilha_exporta(self):
        self.columns = ('data_pedido', 'cliente')
        self.exp_treeview_planilha = ttk.Treeview(self.app11, columns=self.columns, show='headings')

        self.exp_treeview_planilha.heading('data_pedido', text='Data')
        self.exp_treeview_planilha.heading('cliente', text='Cliente')

        self.exp_treeview_planilha.column('data_pedido', width=10)
        self.exp_treeview_planilha.column('cliente', width=80)

        self.exp_treeview_planilha.place(x=910, y=320, width=400, height=240)
        self.scroll_exp_treeview_planilha = Scrollbar(self.app11, orient=VERTICAL,
                                                      command=self.exp_treeview_planilha.yview)

        self.exp_treeview_planilha.configure(yscrollcommand=self.scroll_exp_treeview_planilha.set)
        self.scroll_exp_treeview_planilha.place(x=1313, y=320, width=25, height=240)
        self.exp_treeview_planilha.bind("<Double-1>", self.doubleclickplanilha)

    def buscaclienteplanilha(self):
        self.exp_treeview_planilha.delete(*self.exp_treeview_planilha.get_children())
        self.sql()
        self.nomedecliente= self.nomecliente.get()

        if self.nomedecliente != '':
            self.buscaplanilha=self.cursor.execute(f"""SELECT  DISTINCT data_pedido, cliente FROM planilha
                                    WHERE cliente = '{self.nomedecliente}' """)
            for linha in self.buscaplanilha:
                resultado= linha.data_pedido, linha.cliente
                self.exp_treeview_planilha.insert('','end',values=resultado)
        else:
            self.buscaplanilha = self.cursor.execute(f"""SELECT                              
            DISTINCT TOP 50 data_pedido, cliente FROM planilha """)
            for linha in self.buscaplanilha:
                resultado = linha.data_pedido, linha.cliente
                self.exp_treeview_planilha.insert('', 'end', values=resultado)

    def doubleclickplanilha(self,event):
        self.data_planilha.delete(0, END)
        self.entry_planilha_pesquisa.delete(0, END)
        self.exp_treeview_planilha.selection()
        for y in self.exp_treeview_planilha.selection():
            col1, col2 = self.exp_treeview_planilha.item(y, 'values')
            self.data_planilha.insert(END, col1)
            self.entry_planilha_pesquisa.insert(END, col2)

    #                                        FUNÇAO FATURAMENTO   TELA PRINCIPAL                                      ##

    def faturamento(self):
        self.app12 = Toplevel()
        self.app12.geometry('1350x610+0+56')
        self.app12.resizable(False, False)
        self.app12.transient(self.app)
        self.app12.focus_force()
        self.app12.grab_set()
        self.app12.configure(bg='#2a2a29')


        self.frame_faturamento = Frame(self.app12, bg="white", width=885, height=290, borderwidth=9, relief="raised",
                                    background='#7f7f7f')
        self.frame_faturamento.place(x=10, y=10)

        self.frame_listagem= Frame(self.app12, bg="white", width=885, height=290, borderwidth=9, relief="raised",
                                       background='#7f7f7f')
        self.frame_listagem.place(x=10, y=310)
        self.treviewconsulta()

        self.frame1_faturamento = Frame(self.app12, bg="white", width=450, height=290, borderwidth=9, relief="raised",
                                       background='#7f7f7f')
        self.frame1_faturamento.place(x=900, y=10)

        self.frame1_listagem = Frame(self.app12, bg="white", width=450, height=290, borderwidth=9, relief="raised",
                                        background='#7f7f7f')
        self.frame1_listagem.place(x=900, y=310)
        self.treviewFaturemento()

##                                             CONSULTAR PERIODO

        self.labelFaturamento=Label(self.app12,text='P E R I O D O S  D E  V E N D A S  ', bg='#574c26',
            fg='white', font='times 30', relief='raised')
        self.labelFaturamento.place(x=150,y=30)

        self.inicio=Label(self.app12,text='DATA INICIAL',font='times 15', bg='#7f7f7f',
                                    fg='#ffffff')
        self.inicio.place(x=160,y= 130)

        self.data_inicial = DateEntry(self.app12, dateformat=3, width=15, background='darkblue',
                                      foreground='white', borderwidth=4, locale='pt_BR',
                                      date_pattern='d/m/y', font='TIMES 12 bold')

        self.data_inicial.place(x=20, y=135)

        self.final = Label(self.app12, text='DATA FINAL', font='times 15', bg='#7f7f7f',
                            fg='#ffffff')
        self.final.place(x=160, y=200)

        self.data_final = DateEntry(self.app12, dateformat=3, width=15,
                                    foreground='white', borderwidth=4, locale='pt_BR',
                                    date_pattern='d/m/y', font='TIMES 12 bold')

        self.data_final.place(x=20, y=200)

        self.fotoconsulta=PhotoImage(file='imagens/busca.png')
        self.cons_faturamento=Button(self.app12,text='CONSULTAR',image=self.fotoconsulta,font='times 9',
                                     bg='#7f7f7f',fg='white',command=self.periodo,relief=FLAT)
        self.cons_faturamento.place(x=30, y=230)

        self.imgexel = PhotoImage(file='imagens/exel.png')
        self.exp_consulta = Button(self.app12, text='EXPORTA PERIODO', font='times 10 bold', image=self.imgexel,
                                   bg='#7f7f7f', fg='white',
                                   compound=TOP, relief=FLAT, command=self.exporta_periodo)
        self.exp_consulta.place(x=400, y=220)

        ##                                COMBOBOX                               ##

        self.mes = ["1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12"]

        self.lista_mes = ttk.Combobox(self.app12, width=2, font='times 15 bold', values=self.mes)
        self.lista_mes.configure(foreground='red')
        self.lista_mes.place(x=800, y=100)

        self.lab_consulta=Label(self.app12,text='C O N S U L T A ',relief=FLAT,font='times 15 bold',bg='#7f7f7f',fg='white')
        self.lab_consulta.place(x=550,y=120)

        self.l_combo=Label(self.app12,text='M Ê S ',relief=FLAT,font='times 15 bold',bg='#7f7f7f',fg='yellow')
        self.l_combo.place(x=720,y=100)

        self.l_comboANO = Label(self.app12, text='A N O ', relief=FLAT, font='times 15 bold', bg='#7f7f7f', fg='yellow')
        self.l_comboANO.place(x=720, y=150)

        self.ano = Entry(self.app12,font='times 15', width=5,fg='green')
        self.ano.place(x=800, y=150)

        self.foto=PhotoImage(file='imagens/pesquisadetalhada.png')

        self.btt = Button(self.app12, text='CONSULTAR',image=self.foto,
                          command=self.periodo_mes,bg='#7f7f7f',relief=FLAT)
        self.btt.place(x=800, y=200)


                                     #          COMBOBOX FATURAMENTO #

        self.f_mes = ["1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12"]

        self.lista_f_mes = ttk.Combobox(self.app12, width=2, font='times 15 bold', values=self.mes)
        self.lista_f_mes.configure(foreground='red')
        self.lista_f_mes.place(x=1300, y=100)

        self.lab_consulta_f = Label(self.app12, text='F A T U R A M E N T O ', relief=FLAT, font='times 10 bold', bg='#7f7f7f',
                                  fg='YELLOW')
        self.lab_consulta_f.place(x=1100, y=132)

        self.l_combo_f = Label(self.app12, text='M Ê S ', relief=FLAT, font='times 15 bold', bg='#7f7f7f', fg='yellow')
        self.l_combo_f.place(x=1230, y=100)

        self.l_comboANO_f = Label(self.app12, text='A N O ', relief=FLAT, font='times 15 bold', bg='#7f7f7f', fg='yellow')
        self.l_comboANO_f.place(x=1225, y=150)

        self.ano_f = Entry(self.app12, font='times 15', width=5, fg='green')
        self.ano_f.place(x=1285, y=150)

        self.foto_f = PhotoImage(file='imagens/pesquisadetalhada.png')

        self.btt_f = Button(self.app12, text='CONSULTAR', image=self.foto_f,
                          command=self.periodo_mes, bg='#7f7f7f', relief=FLAT)
        self.btt_f.place(x=800, y=200)

    #              widgets da frame faturamento

        self.lb_faturamento= Label(self.app12,text='F A T U R A M E N T O ',bg='#574c26',
            fg='white', font='times 20', relief='raised')
        self.lb_faturamento.place(x=1000, y=30)

        self.confirmar_faturamento=StringVar()

        self.dia_faturamento = Radiobutton(self.app12,text='DO DIA',value='D',
                                           bg='#7f7f7f',
                                           font='times 12 bold',variable=self.confirmar_faturamento)
        self.dia_faturamento.place(x=910,y=100)

        self.semana_faturamento = Radiobutton(self.app12, text='ULTIMOS 7 DIAS',
                                              value='7',bg='#7f7f7f',
                                              font='times 12 bold',variable=self.confirmar_faturamento)
        self.semana_faturamento.place(x=910,y=130)

        self.quizena_faturamento = Radiobutton(self.app12, text=' ULTIMOS 15 DIAS',
                                           value='15',bg='#7f7f7f',
                                           font='times 12 bold',variable=self.confirmar_faturamento)
        self.quizena_faturamento.place(x=910,y=160)

        self.consulta=Button(self.app12,text='VIZUALIZAR',command=self.lista_faturamento)
        self.consulta.place(x=910,y=200)

        self.consultasoma = Button(self.app12, text='SOMAR FATURAMENTO', command=self.somar_faturamento)
        self.consultasoma.place(x=910, y=260)

        self.consultasoma = Button(self.app12, text='FATURAMENTO MENSAL', command=self.faturamento_total)
        self.consultasoma.place(x=1190, y=180)

        self.consultasoma = Button(self.app12, text='VIZUALIZAR', command=self.vizualizar_faturamento)
        self.consultasoma.place(x=1190, y=210)

   #        PERIODO                      consulta as vendas realizadas entre 2 periodos a data inicial e a data final

    def periodo(self):
        self.sql()
        self.treeview_consulta.delete(*self.treeview_consulta.get_children())
        self.d_inicial=self.data_inicial.get()
        self.d_final= self.data_final.get()

        if self.d_inicial != '':
            tabelass= self.cursor.execute(f""" 
          SELECT data_venda,
          cliente,
          seguimento_venda,
          tema_venda,
          descricao_venda,
          tamanho_venda,
          quantidade_venda,
          valor_venda,
          total_venda FROM seguimento_venda WHERE data_venda between '{self.d_inicial}' AND '{self.d_final}' ORDER BY data_venda
           """)

            for linha in tabelass:

                self.m = linha.data_venda, \
                               linha.cliente, \
                               linha.seguimento_venda, \
                               linha.tema_venda, \
                               linha.descricao_venda, \
                               linha.tamanho_venda, \
                               linha.quantidade_venda, \
                               linha.valor_venda, \
                               linha.total_venda
                self.treeview_consulta.insert('', 'end', values=self.m)
        self.desconectar()

   # PERIODO_MES                      CONSULTA TODAS AS VENDAS REALIZADA NO MES E ANO INFORMADO
    def periodo_mes(self):
        self.sql()
        self.treeview_consulta.delete(*self.treeview_consulta.get_children())
        self.imprimir_mes= self.lista_mes .get()
        self.imprimir_ano=self.ano.get()
        if self.imprimir_mes != '':
            mesConsulta = self.cursor.execute(f""" 
                    SELECT * FROM seguimento_venda WHERE month (data_venda) = '{self.imprimir_mes}' 
                    AND  year (data_venda)= '{self.imprimir_ano}' ORDER BY data_venda
                       """)
            for linha in mesConsulta:
                self.mes = linha.data_venda, \
                            linha.cliente, \
                            linha.seguimento_venda, \
                            linha.tema_venda, \
                            linha.descricao_venda, \
                            linha.tamanho_venda, \
                            linha.quantidade_venda, \
                            linha.valor_venda, \
                            linha.total_venda

                self.treeview_consulta.insert('','end',values=self.mes)
        self.desconectar()

    def faturamento_total(self):
        self.sql()
        self.imprimir_mes_f = self.lista_f_mes.get()
        self.imprimir_ano_f = self.ano_f.get()
        if self.imprimir_mes_f != '':
            mesConsulta = self.cursor.execute(f""" SELECT  SUM(total_venda)as total_venda FROM seguimento_venda 
            WHERE MONTH (data_venda) = '{self.imprimir_mes_f}' AND  YEAR (data_venda)= '{self.imprimir_ano_f}'
                               """)
            for t in mesConsulta:
                fatura=t[0]
                messagebox.showinfo('',f'SEU FATURAMENTO MENSAL REFERENTE AO MÉS '
                                       f'{self.imprimir_mes_f}/{self.imprimir_ano_f}\n '
                                       f' FOI DE \nR$ {fatura:,}')
                # self.fa= linha.total_venda
                # self.treview_faturemento.insert('','end',self.fa)



        self.desconectar()

    def vizualizar_faturamento(self):
        self.sql()
        self.treview_faturemento.delete(*self.treview_faturemento.get_children())
        self.imprimir_mes_f = self.lista_f_mes.get()
        self.imprimir_ano_f = self.ano_f.get()
        if self.imprimir_mes_f != '':
            mesConsulta = self.cursor.execute(
                f""" SELECT data_venda, sum(total_venda)as total_venda from seguimento_venda
                    WHERE month (data_venda) = '{self.imprimir_mes_f}' 
                    AND  year (data_venda)= '{self.imprimir_ano_f}' GROUP BY data_venda""")
            for linha in mesConsulta:
                self.fa=  linha.data_venda,linha.total_venda

                self.treview_faturemento.insert('','end',values=self.fa)

        self.desconectar()

    #  EXPORTA_PERIODO              EXPORTA TODAS AS VENDAS REALIZADA NA DATA INICIAL E NA DATA FINAL

    def exporta_periodo(self):
        import pandas as pd
        import  pyodbc

        self.dt_inicial = self.data_inicial.get()
        self.dt_final = self.data_final.get()

        if self.dt_inicial and self.dt_final !='':
            msg_periodo= messagebox.askyesno('EXPORTA PERIDO', 'SERÁ EXPORTADOS TODOS OS '
                                                               f'SEGUIMENTOS REFERENTE A DATA {self.dt_inicial} '
                                                               f'ATÉ A DATA {self.dt_final} ' 
                                                               'DESEJA CONTINUAR ?')
            if msg_periodo == TRUE:
                self.dados_conexao= ("Driver={SQL Server};"
                                     "Server=DESKTOP-RFPLHG0;"
                                     "Database=seguimentos_cadastrados;")
                self.conexao= pyodbc.connect(self.dados_conexao)
                self.cursor= self.conexao.cursor()

                self.exp_periodo= pd.read_sql(
                                f"""SELECT * FROM seguimento_venda
                                WHERE data_venda BETWEEN '{self.dt_inicial}' AND '{self.dt_final}'""",self.conexao)
                self.expo_perido = pd.DataFrame(data=self.exp_periodo)
                self.ex_vendas=self.exp_periodo[["data_venda","cliente","seguimento_venda","tema_venda",
                                                "descricao_venda","tamanho_venda","quantidade_venda",
                                                "valor_venda","total_venda"]]
                self.ex_vendas.to_excel('periodo.xlsx')
                messagebox.showinfo('CONCLUIDO','EXPORTAÇAO BEM SUCEDIDA')
            else:
                messagebox.showinfo('','NENHUM VALOR SERÁ EXPORTADO')
        else:
            messagebox.showwarning('AVISO','INFORME CORRETAMENTE O PERIDO PARA REALIZAR A EXPORTAÇAO')

    def lista_faturamento(self):
        self.treview_faturemento.delete(*self.treview_faturemento.get_children())

        self.sql()
        self.osvalores = self.confirmar_faturamento.get()

        if self.osvalores == '15':
            quizena=self.cursor.execute("""SELECT data_venda ,SUM(total_venda) as total_venda  FROM seguimento_venda
                                            where (data_venda) between (DATEADD (DAY,-15,GETDATE()))
                                AND (GETDATE()) GROUP BY data_venda """)
            for linha in quizena:
                self.soma= linha.data_venda,linha.total_venda
                self.treview_faturemento.insert('','end',values=self.soma)

        elif self.osvalores == '7':
            quizena=self.cursor.execute("""SELECT data_venda ,SUM(total_venda) as total_venda  FROM seguimento_venda
                                            where (data_venda) between (DATEADD (DAY,-7,GETDATE()))
                                AND (GETDATE()) GROUP BY data_venda """)
            for linha in quizena:
                self.soma= linha.data_venda,linha.total_venda

                self.treview_faturemento.insert('','end',values=self.soma)

    def somar_faturamento(self):
        self.treview_faturemento.delete(*self.treview_faturemento.get_children())

        self.sql()
        self.somaosvalores = self.confirmar_faturamento.get()

        if self.somaosvalores == '15':
            quizena = self.cursor.execute("""SELECT SUM(total_venda)as total_venda FROM seguimento_venda
                                        where (data_venda) between (DATEADD (DAY,-15,GETDATE()))
                                                AND (GETDATE())""")
            for linha in quizena:
                self.soma =  linha.total_venda
                self.treview_faturemento.insert('', 'end', values=self.soma)

        elif self.somaosvalores == '7':
            quizena = self.cursor.execute("""SELECT SUM(total_venda) as total_venda  FROM seguimento_venda
                                                   where (data_venda) between (DATEADD (DAY,-7,GETDATE()))
                                       AND (GETDATE())""")
            for linha in quizena:
                self.soma = linha.data_venda, linha.total_venda

                self.treview_faturemento.insert('', 'end', values=self.soma)



    def  treviewFaturemento(self):
        self.columns=('data_venda','total_venda')

        self.treview_faturemento=ttk.Treeview(self.app12,columns=self.columns,show='headings')
        self.treview_faturemento.heading('data_venda', text='DATA')
        self.treview_faturemento.heading('total_venda', text='FATURAMENTO')

        self.treview_faturemento.column('data_venda', width=30)
        self.treview_faturemento.column('total_venda', width=30)

        self.treview_faturemento.place(x=910,y=320,width=430,height=270)

    #                              TREEVIEW DE CONSULTA

    def treviewconsulta(self):
        self.columns = (
            'data', 'cliente', 'seguimento', 'tema', 'descricao', 'tamanho', 'quantidade', 'valor', 'total_venda')
        self.treeview_consulta = ttk.Treeview(self.app12, columns=self.columns, show='headings')
        self.treeview_consulta.heading('data', text='Data')
        self.treeview_consulta.heading('cliente', text='Cliente')
        self.treeview_consulta.heading('seguimento', text='Seguimento')
        self.treeview_consulta.heading('tema', text='Tema')
        self.treeview_consulta.heading('descricao', text='Descriçao')
        self.treeview_consulta.heading('tamanho', text='Tamanho')
        self.treeview_consulta.heading('quantidade', text='Quantidade')
        self.treeview_consulta.heading('valor', text='Valor')
        self.treeview_consulta.heading('total_venda', text='Total')
        self.treeview_consulta.column('cliente', width=50)
        self.treeview_consulta.column('data', width=30)
        self.treeview_consulta.column('seguimento', width=170)
        self.treeview_consulta.column('tema', width=50)
        self.treeview_consulta.column('descricao', width=110)
        self.treeview_consulta.column('tamanho', width=40)
        self.treeview_consulta.column('quantidade', width=22)
        self.treeview_consulta.column('valor', width=5)
        self.treeview_consulta.column('total_venda', width=6)
        self.treeview_consulta.place(x=20, y=320,width=850,height=270)
        self.scroll_consul = Scrollbar(self.app12, orient=VERTICAL, command=self.treeview_consulta.yview)
        self.treeview_consulta.configure(yscrollcommand=self.scroll_consul.set)
        self.scroll_consul.place(x=873, y=320, width=15, height=270)


Funcoes()