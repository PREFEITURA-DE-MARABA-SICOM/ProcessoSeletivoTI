# -*- coding: utf-8 -*-  
import pandas as pd

class Servidor:
    
    def __init__(self,nomeServidor, matriculaServidor, file):

        self.nomeServidor=nomeServidor ;self.matriculaServidor=matriculaServidor ;self.file=file ;self.df=pd.read_excel(self.file, index_col=None, dtype={'Nº do Registro':str, 'Nome':str, 'Data':str, 'Atividades':str, 'Endereço':str, 'Bairro':str, 'Servidor':str, 'Procedimentos':str}) ;self.df = self.df.dropna()
    
    def PullDf(self):
        return self.df
    
    def InsertLine(self, line):
        self.df = self.df.append(line, ignore_index=True).reset_index()
    
    def UpdateLine(self, line, col, val):
        self.df = self.df.replace(to_replace=self.df.loc[line,col], value=val)
        
    def DeleteLine(self, colIndex):
        self.df = self.df.drop(colIndex).reset_index() 


s1 = Servidor('ServidorTeste','01','RelatorioPS.xlsx')


#insercao de dados-----------------------------------------------
numRegistro = '800' ; nome = 'Servidor aleatorio' ; data = '16/10/2020' ; atividade = 'PRESTACAO DE SERVICOS' ; endereco = 'RUA RUI BARBOSA' ; bairro = 'BOM PLANALTO' ; servidor = 'Servidor teste' ; procedimentos = 'Relatorio'
line = {'Nº do Registro':numRegistro, 'Nome':nome, 'Data':data, 'Atividades':atividade, 'Endereço':endereco, 'Bairro':bairro, 'Servidor':servidor,'Procedimentos':procedimentos}
s1.InsertLine(line)

#atualizacao de dados--------------------------------------------
s1.UpdateLine(78, 'Procedimentos', 'emissao nota fiscal') #alterando a coluna 'Procedimentos'

#exclusao de dados-----------------------------------------------
s1.DeleteLine(77)

#salvando novo bd
excel = pd.ExcelWriter('NovoBD.xlsx', engine='xlsxwriter'); s1.df.to_excel(excel, sheet_name='novoBD'); excel.save()


