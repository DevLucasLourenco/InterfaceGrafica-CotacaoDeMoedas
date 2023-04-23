import numpy as np
from datetime import datetime
import tkinter as tk 
from tkinter import ttk
from tkcalendar import DateEntry
import requests
import json
import locale
from tkinter import filedialog
import pandas as pd



class CotacoesMoeda:
    
    def __init__(self):
        
        self.requisicao_json()
        self.janela_grafica()


    @staticmethod
    def formatar_moeda(valor):
        locale.setlocale(locale.LC_MONETARY, '')  
        return locale.currency(valor, grouping=True)


    def requisicao_json(self):
        requisicao = requests.get('https://economia.awesomeapi.com.br/json/all')
        dicionario_moedas = requisicao.json()
        self.lista_moedas = list(dicionario_moedas.keys())


    def pegar_cotacao(self):
        moeda = self.combobox_selecionar_moeda.get()
        data_cotacao = self.calendario_moeda.get()
        dia, mes, ano = data_cotacao.split('/')
        data_cotacao = f'{ano}{mes}{dia}'
        
        try:
            requisicao_moeda = requests.get(f'https://economia.awesomeapi.com.br/{moeda}-BRL/10?start_date={data_cotacao}&end_date={data_cotacao}').json()
            valor_moeda = requisicao_moeda[0]['bid']
            self.label_texto_cotacao['text'] = f'A cotação da moeda {moeda} no dia {dia}/{mes}/{ano} foi de {CotacoesMoeda.formatar_moeda(float(valor_moeda))}'
        
        except Exception:
            self.label_texto_cotacao['text'] = f'Não foi possível encontrar a cotação da moeda {moeda} \nno dia {dia}/{mes}'
           
            
    def selecionar_arquivo(self):
        self.caminho_arquivo = filedialog.askopenfilename(title='Escolha o arquivo .xlsx de moeda.')
        self.nome_arquivo = self.caminho_arquivo.split('/')[-1]
        
        if self.caminho_arquivo:
            self.label_arquivo_selecionado['text'] = self.nome_arquivo


    def atualizar_cotacoes(self):
        try:
            df = pd.read_excel(self.caminho_arquivo)
            moedas = list(df.iloc[:, 0])
            
            data_inicial = self.calendario_data_inicial.get()
            data_final = self.calendario_data_final.get()
            
            ano_inicial = data_inicial[-4:]
            mes_inicial = data_inicial[3:5]
            dia_inicial = data_inicial[:2]

            ano_final = data_final[-4:]
            mes_final = data_final[3:5]
            dia_final = data_final[:2]
            
            quantidade = 200
            novo_df = pd.DataFrame()
            
            for moeda in moedas:
                novo_df[moeda] = np.nan
                
                cotacoes = requests.get(f'https://economia.awesomeapi.com.br/json/daily/{moeda}-BRL/{quantidade}?' \
                    f'start_date={ano_inicial}{mes_inicial}{dia_inicial}&' \
                    f'end_date={ano_final}{mes_final}{dia_final}').json()
                
                for cotacao in cotacoes:
                    
                    timestamp = int(cotacao['timestamp'])
                    data = datetime.fromtimestamp(timestamp)
                    data = data.strftime('%d/%m/%Y')
                    
                    bid = float(cotacao['bid'])
                    novo_df.loc[data, moeda] = bid

            novo_df.insert(0, column='Data', value=novo_df.index)
            novo_df.index = pd.to_datetime(novo_df.index, dayfirst=True)
            novo_df.sort_index(inplace=True)
            novo_df.to_excel(f'Moedas - {dia_inicial}.{mes_inicial}.{ano_inicial} - {dia_final}.{mes_final}.{ano_final}.xlsx', index=False)
            self.label_atualizarcotacoes['text'] = 'Arquivo atualizado com sucesso'
        except:
            self.label_atualizarcotacoes['text'] = 'Selecione o formato de arquivo correto.'
            
        
            

    def janela_grafica(self):

        root = tk.Tk()
        root.title('Janela Gráfica')


        label_cotacao_moeda = tk.Label(text='Cotação de 1 moeda específica', borderwidth=2, relief='solid')
        label_cotacao_moeda.grid(row=0, column=0, padx=10, pady=10, columnspan=3, sticky='nswe')


        label_selecionar_moeda = tk.Label(text='Selecionar Moeda', anchor='e')
        label_selecionar_moeda.grid(row=1, column=0, padx=10, pady=10, columnspan=2)

        self.combobox_selecionar_moeda = ttk.Combobox(values=self.lista_moedas)
        self.combobox_selecionar_moeda.grid(row=1, column=2, padx=10, pady=10, sticky='nswe') 


        label_selecionar_dia = tk.Label(text='Selecione o dia que deseja pegar a cotação', anchor='e')
        label_selecionar_dia.grid(row=2, column=0, padx=10, pady=10, columnspan=2)

        self.calendario_moeda = DateEntry(year=2023, locale='pt_br')
        self.calendario_moeda.grid(row=2, column=2, padx=10, pady=10, sticky='nswe')

        self.label_texto_cotacao = tk.Label(text='')
        self.label_texto_cotacao.grid(row=3, column=0, padx=10, pady=10, sticky='nswe')


        self.label_pegar_cotacao = tk.Button(text='Pegar Cotação',command=self.pegar_cotacao)
        self.label_pegar_cotacao.grid(row=3, column=2, padx=10, pady=10, sticky='nswe')



        # cotacao de várias moedas

        label_cotacao_multiplas_moedas = tk.Label(text='Cotação de múltiplas moedas', borderwidth=2, relief='solid')
        label_cotacao_multiplas_moedas.grid(row=4, column=0, padx=10, pady=10, columnspan=3, sticky='nswe')


        label_selecionar_arquivo = tk.Label(text='Selecione um arquivo em excel com as moedas na Coluna A')
        label_selecionar_arquivo.grid(row=5, column=0, padx=10, pady=10, columnspan=2, sticky='nswe')

        botao_selecionar_arquivo = tk.Button(text='Clique para Selecionar', command=self.selecionar_arquivo)
        botao_selecionar_arquivo.grid(row=5, column=2, padx=10, pady=10, sticky='nswe')


        self.label_arquivo_selecionado = tk.Label(text='Nenhum Arquivo Selecionado', anchor='e')
        self.label_arquivo_selecionado.grid(row=6, column=0, padx=10, pady=10, columnspan=3, sticky='nswe')


        label_data_inicial = tk.Label(text='Data Inicial')
        label_data_inicial.grid(row=7, column=0, padx=10, pady=10, sticky='nswe')

        self.calendario_data_inicial = DateEntry(year=2023, locale='pt_br')
        self.calendario_data_inicial.grid(row=7, column=1, padx=10, pady=10, sticky='nswe')


        label_data_final = tk.Label(text='Data Final')
        label_data_final.grid(row=8, column=0, padx=10, pady=10, sticky='nswe')

        self.calendario_data_final = DateEntry(year=2023, locale='pt_br')
        self.calendario_data_final.grid(row=8, column=1, padx=10, pady=10, sticky='nswe')


        botao_atualizar_cotacoes = tk.Button(text='Atualizar Cotacoes', command=self.atualizar_cotacoes)
        botao_atualizar_cotacoes.grid(row=9, column=0, padx=10, pady=10, sticky='nswe')

        self.label_atualizarcotacoes = tk.Label(text="")
        self.label_atualizarcotacoes.grid(row=9, column=2, columnspan=2, padx=10, pady=10, sticky='nswe')

        botao_fechar = tk.Button(text='Fechar', command=root.quit)
        botao_fechar.grid(row=10, column=2, padx=10, pady=10, sticky='nswe')


        root.mainloop()
        
        
if __name__ == '__main__':
    instancia = CotacoesMoeda()