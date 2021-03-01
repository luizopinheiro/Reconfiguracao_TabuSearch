
import numpy as np
import win32com.client
from datetime import datetime
import random

def DecodificadorChaves (Individuo, NomesChaves):
    ChavesAbertas = []
    for i in range(len(NomesChaves)):
        if (Individuo[i]==0):
            ChavesAbertas.append(NomesChaves[i])

    return ChavesAbertas


def FitnessIndividuo(S, NomesChaves):
    
    for i in range(len(S)):
        objeto.switch_elemento(NomesChaves[i], S[i])
        
    objeto.solve_DSS_snapshot()       
    return objeto.get_line_losses()

def Vizinhos (S, N_Malhas, N_Chaves_Malha):
    
    Vizinhos = []
    inicio = 0
    changes_list = [] #lista dos índices que foram alterados (para lista tabu)
    Vizinhos2 = []
    Vizinhos3=[]


    #atualmente estou escolhendo um aleatorio em cada malha. no primeiro vizinho esse aleatorio vai pra 0 
    #e onde tá 0 vai pra 1
    #nos demais vizinhos vai andando com o 0 pra frente de onde estava no primeiro vizinho
    i = 0
    for i in range(N_Malhas): #for dos vizinhos
        if (i==0):
            Vizinho = S
        else:
            Vizinho = Vizinhos[i-1]
        inicio = 0

        for j in range (N_Malhas): #for das malhas  
            if (i==0): #primeiro vizinho gera o aleatorio e substitui
                random.seed(datetime.now())
                x = random.randint(inicio, inicio + N_Chaves_Malha[j]-1)
                index = Vizinho.index(0, inicio, inicio + N_Chaves_Malha[j])
                if (Vizinho[x]==1):
                    Vizinho[x] = 0
                    Vizinho[index] = 1
                
                changes_list.append(x)
                inicio = inicio + N_Chaves_Malha[j]
                
                
            else: #demais vizinhos pega o vizinho anterior e só anda com o 0 pra frente
                index = Vizinho.index(0, inicio, inicio + N_Chaves_Malha[j])
                if (index == (inicio + N_Chaves_Malha[j]-1)):         
                    Vizinho[index] = 1
                    Vizinho[inicio] = 0
                    changes_list.append(inicio)
                else:
                    Vizinho[index] = 1
                    Vizinho[index+1] = 0
                    changes_list.append(index+1)            
                
                inicio = inicio + N_Chaves_Malha[j]

             
        Vizinhos2.append(np.array(Vizinho))
        Vizinhos.append(Vizinho)
        
    for i in range(len(Vizinhos2)):
        Vizinhos3.append(Vizinhos2[i].tolist())
        
                                      
    return Vizinhos3, changes_list
        

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
    
    #NomesChaves=["S1","S2","S3","S4","S5","S6","S7","S8","S9","S10","S11","S12","S13","S14","S15"]
    #NomesChaves=["S1","S2","S3","S4","S5","S6","S7"]
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
    BTMax = 1

    print ("Fit inicial: ", FitS)
    
    i = 0
    BestIter = 1
    T=10
    
    FitVizS = []
    lista_tabu = []
    ChavesAbertasVizS = []
    
    while (i <= 20):
    #while (i-BestIter) <= BTMax:
        i+=1
        print ("================NOVA ITERAÇÃO=================", i)
        print ("Partindo com Solução Corrente:", DecodificadorChaves(S, NomesChaves))
        VizS, lista_movimentos = Vizinhos(S, N_Malhas, N_Chaves_Malha)
        #lista_movimentos é um array que guarda todos os movimentos feitos em cada malha em cada vizinho
        #ou seja, lista_movimentos = [0, 1, 2, 4, 5, 6], por exemplo, 0 1 2 foram os movimentos para as malhas 1, 2 e 3 do primeiro vizinhi
        #enquanto 4, 5 6 foram movimentos das malhas 1 2 e 3 para o segundo vizinho etc

        #gera os fits dos vizinhos e a lista das chaves abertas em cada vizinho
        for j in range(len(VizS)):
            FitVizS.append(FitnessIndividuo(VizS[j], NomesChaves))
            ChavesAbertasVizS.append(DecodificadorChaves(VizS[j], NomesChaves))
            
        FitSorted = sorted(FitVizS, reverse=False) #organiza os Fit em ordem decrescente
        index_melhorfit = FitVizS.index(FitSorted[0]) #obtém o index do maior fit no array original       
        S = VizS[index_melhorfit] #nova solução corrente 
        
        lista_tabu.clear() 
        #obtém somente a lista de movimentos que foram realizados naquele indivíduo, em cada malha
        for j in range(N_Malhas*index_melhorfit, N_Malhas*index_melhorfit + N_Malhas):
            lista_tabu.append(lista_movimentos[j])
            
        for j in range(len(VizS)):
            #print ("Vizinho: " ,VizS[j])
            print ("Vizinho: ", ChavesAbertasVizS[j])
            print ("Fit Vizinho: ", FitVizS[j])
            
          
        print ("S escolhido", DecodificadorChaves(S, NomesChaves))
        #print ("Lista movimentos: ", lista_movimentos)
        
        if (FitVizS[index_melhorfit] < BestFit):
            BestFit = FitVizS[index_melhorfit]
            BestS = VizS[index_melhorfit]
        
        print ("MELHOR S ATÉ O MOMENTO: ", DecodificadorChaves(BestS, NomesChaves))
        print ("MELHOR FIT ATÉ O MOMENTO: ", BestFit)
        print("lista_tabu", lista_tabu) 
        FitVizS.clear()
        VizS.clear()
        FitSorted.clear()
        ChavesAbertasVizS.clear()
        lista_movimentos.clear()
        
            
        
        

    





    
    
