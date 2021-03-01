import numpy as np
import win32com.client
from datetime import datetime
import random
import time 

start_time = time.time()

def FitnessIndividuo(S, NomesChaves):
    
    for i in range(len(S)):
        objeto.switch_elemento(NomesChaves[i], S[i])
        
    objeto.solve_DSS_snapshot()       
    return objeto.get_line_losses()

def DecodificadorChaves (Individuo, NomesChaves):
    ChavesAbertas = []
    for i in range(len(NomesChaves)):
        if (Individuo[i]==0):
            ChavesAbertas.append(NomesChaves[i])

    return ChavesAbertas

class DSS():

    def __init__(self, end_modelo_DSS):

        self.end_modelo_DSS = end_modelo_DSS

        # Criar a conexão entre Python e OpenDSS
        self.dssObj = win32com.client.Dispatch("OpenDSSEngine.DSS")

        # Iniciar o Objeto DSS
        if self.dssObj.Start(0) == False:
            print ("Problemas em iniciar o OpenDSS")
        else:
            # Criar variáveis paras as principais interfaces
            self.dssText = self.dssObj.Text
            self.dssCircuit = self.dssObj.ActiveCircuit
            self.dssSolution = self.dssCircuit.Solution
            self.dssCktElement = self.dssCircuit.ActiveCktElement
            self.dssBus = self.dssCircuit.ActiveBus
            self.dssLines = self.dssCircuit.Lines
            self.dssTransformers = self.dssCircuit.Transformers
            #self.dssEngine = self.dssObj.
            #self.dssSwtControl = self.dssObj.SwtControls

    def versao_DSS(self):

        return self.dssObj.Version

    def compile_DSS(self):

        # Limpar informações da ultima simulação
        self.dssObj.ClearAll()

        self.dssText.Command = "compile " + self.end_modelo_DSS

    def solve_DSS_snapshot(self):

        # Configurações
        #self.dssText.Command = "Set Mode=Snap"
        #self.dssText.Command = "Set ControlMode=STATIC"

        # Resolve o Fluxo de Potência
        self.dssSolution.Solve()

    def get_resultados_potencia(self):
        self.dssText.Command = "Show powers kva elements"

    def get_nome_circuit(self):
        return self.dssCircuit.Name
    
    def switch_elemento(self,chave,estado):
        if (estado==1 or estado=='1'):
            self.dssText.Command = "SwtControl."+chave+".action=close"
        elif (estado==0 or estado=='0'):
            self.dssText.Command = "SwtControl."+chave+".action=open"
            
    def matrixY(self):
        self.dssText.Command = "build Y"
        self.dssText.Command = "calcincmatrix"
        self.dssText.Command = "show Y"
    
    def get_line_losses (self):
        perdas,perdas2 = self.dssCircuit.LineLosses
        return perdas
    
if __name__ == "__main__":
    print ("""Autor: Luiz Otávio""")

    # Criar um objeto da classe DSS
    # Inserir aqui o nome da pasta onde encontra-se o seu .dss
    # Caso haja espaços no endereço, colocar entre []
    objeto = DSS(r"[C:\Users\luiz3\Google Drive (luiz.filho@cear.ufpb.br)\Tese\Código\sistemas\33_barras.dss]")
    objeto.compile_DSS()
    print (u"Versão do OpenDSS: " + objeto.versao_DSS() + "\n")
    
    NomesChaves=["S2","S3","S4","S5","S6","S7","S33","S20","S19","S18",
                 "S22","S23","S24","S37","S28","S27","S26","S25",
                 "S8","S9","S10","S35","S21",
                 "S29","S30","S31","S32","S36","S17","S16","S15",
                 "S34","S14","S13","S12","S11"]
    N_Malhas = 5
    N_Chaves_Malha = []
    N_Chaves_Malha.append(10) #numero de chaves da malha 1
    N_Chaves_Malha.append(8) #numero de chaves da malha 2
    N_Chaves_Malha.append(5) #numero de chaves da malha 3
    N_Chaves_Malha.append(8) #numero de chaves da malha 3
    N_Chaves_Malha.append(5) #numero de chaves da malha 3
    
    S = [1,1,1,1,1,1,0,1,1,1,
         1,1,1,0,1,1,1,1,
         1,1,1,0,1,
         1,1,1,1,0,1,1,1,
         0,1,1,1,1]

    FitS = FitnessIndividuo(S, NomesChaves)
    print ("Configuração inicial: ", DecodificadorChaves(S, NomesChaves))
    
    BestFit = FitS
    BestS = S
    Iter = 0
    BTMax = 20

    print ("Fit inicial: ", FitS)
    
    i = 0 #contador de iterações
    BestIter = 1 #iteração (i) que obteve o melhor Fit
    T = 4 #tamanho da lista tabu
    
    FitVizS = []
    lista_tabu = [-1]*sum(N_Chaves_Malha) #declaração da lista tabu
    novo_tabu = [-1]*sum(N_Chaves_Malha) #declaração de novo_tabu
    ChavesAbertasVizS = []
    lista_tabu_completa = []
    contador_novo_S = 0
    
    while (i-BestIter) <= BTMax:
        i+=1
        print ("================NOVA ITERAÇÃO=================", i)
        print ("Partindo com Solução Corrente:", DecodificadorChaves(S, NomesChaves))
        VizS, lista_movimentos, lista_movimentos_separados = Vizinhos(S, N_Malhas, N_Chaves_Malha)
