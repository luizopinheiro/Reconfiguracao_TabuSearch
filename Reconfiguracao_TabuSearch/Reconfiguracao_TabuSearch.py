"""Especifique quantas vezes você quer que o programa rode:"""
N_Vezes = 30

"""Especifique o endereço do arquivo Opendss (dentro dos colchetes):"""
address_opendss = r"[C:\Users\luiz3\Google Drive (luiz.filho@cear.ufpb.br)\Tese\Código\sistemas\69_barras.dss]"

"""Especifique o número de barras do sistema que você quer trabalhar (33, 69 ou 94):"""
N_Barras_Sistema = 69

"""Especifique quantas iterações sem melhora o sistema deve fazer no máximo"""
BTMax = 200

"""Especifique o tamanho da Lista Tabu"""
T = 3

"""Número máximo de iterações"""
MaxIter = 2000

"""Versão do código"""
versao_codigo = "12.3"

MelhoresFitsGerais = []
MelhoresSGerais = []
MinimasTensoesGerais = []
MaximasTensoesGerais = []
TempoGasto = []

solucoesotimas = 0
k = 0

from openpyxl import load_workbook
wb = load_workbook(r"C:\Users\luiz3\Google Drive (luiz.filho@cear.ufpb.br)\Tese\Resultados_TS.xlsx")
ws = wb.worksheets[0]

for i in range (N_Vezes):
#while k < N_Vezes:
    
    import numpy as np
    import win32com.client
    import random 
    import copy
    import time
   
    start = time.time()
    
    def S_inicial_PFO(NomesTodasChaves): 
    
        objeto.solve_DSS_snapshot()       
        potencias_ativas_trifasicas = []
        potencias_malha_aux = []
        
        NomesTodasChaves_aux = copy.deepcopy(NomesTodasChaves) 
                
        #essa parte obtém o fluxo de cada uma das linhas e armazena no vetor potencias_ativas_trifasicas
        #é válido comentar que esse vetor corresponde com o NomesChaves
        for i in range (len(NomesTodasChaves)): #for de cada uma das malhas
            potencias_malha_aux = []
            for j in range (len(NomesTodasChaves[i])): #for das chaves dentro de cada malha
                objeto.dssSwtControl.Name = NomesTodasChaves_aux[i][j] #ativa o switch pelo nome da chave
                switched_obj = objeto.dssSwtControl.SwitchedObj  #descobre o nome da linha referente
                objeto.ativa_elemento(switched_obj) #ativa a determinada linha
                b1, b2 = objeto.get_barras_elemento()
                potencias = (objeto.get_potencias_elemento()) #pega todas as potencias da linha
                potencias_malha_aux.append(abs(potencias[0] + potencias[2] + potencias[4])) #soma somente as potências ativas entrando na linha
            potencias_ativas_trifasicas.append(potencias_malha_aux)
            
            
        
        #for i in range (len(NomesTodasChaves)):
            #print ("=====Malha", i,"======")
            #for j in range (len(NomesTodasChaves[i])):
            #    print ("Potencia na chave", NomesTodasChaves[i][j],": ", potencias_ativas_trifasicas[i][j])
        
        chaves = []
        for i in range (len(potencias_ativas_trifasicas)): #for de cada uma das malhas
            potencias_sorted = sorted(potencias_ativas_trifasicas[i], reverse=False) #ordena os fluxos de potência de forma crescente para uma determinada malha
            menor_fluxo_chave_malha = potencias_sorted[0]
            index_menor_fluxo_chave_malha = potencias_ativas_trifasicas[i].index(menor_fluxo_chave_malha) #obtém o index do menor fluxo no vetor original potencias_ativas_trifasicas
            chave = NomesTodasChaves[i][index_menor_fluxo_chave_malha] #obtém a chave no vetor NomesTodasChaves
            #print ("====Malha", i,"====")
            #print ("Chave inicialmente escolhida:", chave)
            c = 0
            k = 0

            #while c < len(chaves):
            for c in range(len(chaves)):
                ChavesComuns = list(set(NomesTodasChaves_aux[i]) & set(NomesTodasChaves_aux[c])) #chaves comuns entre a malha M e a malha C (das chaves anteriores)
                #print ("ChavesComuns entre as malhas", i,"e", c,": ", ChavesComuns)
                if ((chave in ChavesComuns) and (chaves[c] in ChavesComuns)):
                 #   print ("A chave escolhida", chave, "e a chave", chaves[c], "estão nas chaves comuns")
                  #  print ("NomesTodasChaves[",c,"]=",NomesTodasChaves[c])
                   # print ("NomesTodasChaves[",i,"]=",NomesTodasChaves[i])
                    ChavesNaoComunsMalhaM = [x for x in NomesTodasChaves_aux[i] if x not in ChavesComuns] #obtém as chaves que não são comuns entre a malha I e a malha da chave em comum
                    potencias_aux = []
                    #print ("ChavesNaoComuns entre as malhas",i,"e",c,ChavesNaoComunsMalhaM)
                    
                    for j in range (len(ChavesNaoComunsMalhaM)): 
                        index_aux = NomesTodasChaves_aux[i].index(ChavesNaoComunsMalhaM[j]) #encontra o index das chaves não comuns dentro do vetor NomesChaves
                        potencias_aux.append(potencias_ativas_trifasicas[i][index_aux]) #adiciona ao vetor potencias_aux os fluxos das chaves não comuns
                    potencias_aux_sorted = sorted(potencias_aux, reverse=False)
                
                    menor_fluxo_chave_malha = potencias_aux_sorted[0] #pega o menor fluxo entre as chaves não comuns
                    index_menor_fluxo_chave_malha = potencias_ativas_trifasicas[i].index(menor_fluxo_chave_malha)
                    chave = NomesTodasChaves_aux[i][index_menor_fluxo_chave_malha] #obtém no vetor NomesChaves a chave de menor fluxo
                    #print ("Chave finalmente escolhida: ", chave)
                    break
            
            chaves.append(chave)
        
        print ("S inicial: ", chaves) 
        
        return chaves
    
    def Ajuste_SistemaNBarras (N_Barras):
        
        if (N_Barras == 69):
            NomesTodasChaves = [0,0,0,0,0]
            NomesTodasChaves[0] = ["S4","S5","S6","S7","S8","S46","S47","S48","S49","S52","S53",
                          "S54","S55","S56","S57","S58","S72"]
            NomesTodasChaves[1] = ["S3","S4","S5","S6","S7","S8","S9","S10","S35","S36","S37",
                          "S38","S39","S40","S41","S42","S69"]
            NomesTodasChaves[2] = ["S11","S12","S13","S14","S43","S44","S45","S69","S71"]
            NomesTodasChaves[3] = ["S13","S14","S15","S16","S17","S18","S19","S20","S70"]
            NomesTodasChaves[4] = ["S9","S10","S11","S12","S21","S22","S23","S24","S25","S26","S52","S53",
                      "S54","S55","S56","S57","S58","S59","S60","S61","S62","S63",
                      "S64","S73","S70"]
            return NomesTodasChaves
        elif (N_Barras == 33):
            NomesTodasChaves = [0,0,0,0,0]
            NomesTodasChaves[0] = ["S2", "S3","S4","S5","S6","S7","S33","S20","S19","S18"]
            NomesTodasChaves[1] = ["S22","S23","S24","S37","S28","S27","S26","S25","S5","S4","S3"]
            NomesTodasChaves[2] = ["S33","S8","S9","S10","S11","S35","S21"]
            NomesTodasChaves[3] = ["S25","S26","S27","S28","S29","S30","S31","S32","S36","S17","S16","S15","S34","S8","S7","S6"]
            NomesTodasChaves[4] = ["S34","S14","S13","S12","S11","S10","S9"]
            return NomesTodasChaves
        elif (N_Barras == 94):
            NomesTodasChaves = [0]*13
            NomesTodasChaves[0] = ["S84","S54","S55","S96","S63","S62","S61","S85","S6",
                                   "S7","S64"]
            NomesTodasChaves[1] = ["S87","S72","S70","S69","S68","S67","S66","S65","S13",
                                   "S88","S75","S74","S73","S76","S71"]
            NomesTodasChaves[2] = ["S77","S78","S79","S80","S81","S82","S83","S91","S20",
                                   "S19","S89","S14","S13","S88","S76","S75","S74","S73"]
            NomesTodasChaves[3] = ["S65","S66","S67","S68","S69","S70","S71","S72","S87",
                                   "S11","S12"]
            NomesTodasChaves[4] = ["S12","S14","S89","S18","S17","S90","S27","S28","S29",
                                   "S93","S40","S95","S94","S86","S35","S36","S37","S38",
                                   "S41","S42","S44","S45","S46"]
            NomesTodasChaves[5] = ["S43","S44","S45","S46","S94","S30","S31","S32","S33",
                                   "S34"]
            NomesTodasChaves[6] = ["S41","S42","S95","S40","S39"]
            NomesTodasChaves[7] = ["S33","S34","S35","S36","S37","S38","S39","S93","S29",
                                   "S92"]
            NomesTodasChaves[8] = ["S30","S31","S32","S92","S25","S26","S27","S28"]
            NomesTodasChaves[9] = ["S25","S26","S90","S15","S16"]
            NomesTodasChaves[10] = ["S47","S48","S49","S50","S51","S52","S53","S1","S2",
                                    "S3","S4","S5","S84","S54","S55"]
            NomesTodasChaves[11] = ["S1","S2","S3","S4","S5","S6","S7","S85","S56","S57",
                                    "S58","S59","S60"]
            NomesTodasChaves[12] = ["S11","S86","S43"]
            return NomesTodasChaves
            
    def Tensao_Sistema (N_Barras):
        if (N_Barras == 33):
            return 7309.25
        elif (N_Barras == 69):
            return 7309.25
        elif (N_Barras == 94):
            return 6581.8
            
    #esta função organiza as chaves levando em consideração uma solução S.
    #desta maneira as malhas/chaves são reorganizadas obrigando cada chave aberta de S a pertencer a apenas uma malha
    #desta maneira, após a reorganização cada malha terá uma chave aberta (um 0)
    def OrganizaChaves_com_S_ch (NomesTodasChaves, S_ch):
        
        N_Chaves_Malha = []
        NomesTodasChaves_aux = copy.deepcopy(NomesTodasChaves)

        for i in range (len(S_ch)):
            for j in range (len(NomesTodasChaves_aux)):
                if (i!=j):
                    if S_ch[i] in NomesTodasChaves_aux[j]:
                        NomesTodasChaves_aux[j].remove(S_ch[i])
        
        for i in range (len(NomesTodasChaves_aux)):
            for j in range (len(NomesTodasChaves_aux)):
                if (i!=j):
                    ChavesComuns = list(set(NomesTodasChaves_aux[i]) & set(NomesTodasChaves_aux[j]))
                    #print ("ChavesComuns", ChavesComuns)
                    if ChavesComuns: 
                        for k in range (len(ChavesComuns)):
                            Aleatorio = np.random.randint(2) 
                            if (Aleatorio == 0): #se 0 a chave fica na malha i, se 1 fica na malha j
                                NomesTodasChaves_aux[j].remove(ChavesComuns[k])
                            elif (Aleatorio == 1):
                                NomesTodasChaves_aux[i].remove(ChavesComuns[k])
                        ChavesComuns.clear()
                        ChavesComuns = []
        
        for i in range (len(NomesTodasChaves_aux)):
            N_Chaves_Malha.append(len(NomesTodasChaves_aux[i]))
        
        N_Malhas = len(NomesTodasChaves_aux)
        
        NomesChaves_aux = sum(NomesTodasChaves_aux, [])
        
        return NomesChaves_aux, N_Chaves_Malha, N_Malhas
    
        
    #esta função organiza o Fit em ordem crescente e consequentemente os Vizinhos e a lista de movimentos 
    #em ordem crescente de acordo com o Fit já ordenado 
    def SortFitVizS_VizS_ListaMovimentos(FitVizS, N_Malhas, VizS_ch):
        
        Fit_Sorted = sorted(FitVizS, reverse=False)
        VizS_ch_Sorted = []
        
        for i in range (len(FitVizS)):
            index = FitVizS.index(Fit_Sorted[i])
            VizS_ch_Sorted.append(VizS_ch[index])
            
        return Fit_Sorted, VizS_ch_Sorted
                        
    #Esta função recebe um S qualquer (binário) e transforma em chaves S1, S2 etc.
    def DecodificadorChaves (Individuo, NomesChaves):
        ChavesAbertas = []
        for i in range(len(NomesChaves)):
            if (Individuo[i]==0):
                ChavesAbertas.append(NomesChaves[i])
    
        return ChavesAbertas

    #esta função recebe um S qualquer (S1, S2 etc) e transforma em binário 0,1,0,1...
    def CodificadorChaves(ChavesAbertas, NomesChaves):
        
        Individuo = [1]*len(NomesChaves)
        for i in range(len(ChavesAbertas)):
            index = NomesChaves.index(ChavesAbertas[i])
            Individuo[index] = 0
        
        return Individuo
    
    #calcula o fitness de um S chaves qualquer.    
    def FitnessIndividuo_ch(Ind_ch):
        for i in range (len(NomesChaves)):
            objeto.switch_elemento(NomesChaves[i], 1)
        
        for i in range(len(Ind_ch)):
            objeto.switch_elemento(Ind_ch[i], 0)
            
        objeto.solve_DSS_snapshot()       
        return objeto.get_line_losses()
    
    
    #esta função gera vizinhos malha a malha, ou seja, para cada malha numa solução S, a função sorteia um lado (esq ou dir)
    #para "andar" com a solução. 
    #ela é capaz de gerar não somente um vizinho, mas mais de um, ou seja, os dois ou três da esquerda etc
    def Vizinhos_LR_N (S_ch, N_Malhas, N_Chaves_Malha, N_Barras_Sistema):
        
        if (N_Barras_Sistema == 33) or (N_Barras_Sistema == 94):
            N = 1
        else:
            N = 3
        
        NomesChaves_aux = NomesChaves.copy()
               
        Vizinhos = []
        FitVizinhos = []
        
        N_Chaves_Malha_index = []
        soma = 0
    
        for i in range (len(N_Chaves_Malha)):
            soma = N_Chaves_Malha[i] + soma
            N_Chaves_Malha_index.append(soma - 1)
        
        for i in range (1,N+1):
            
            for v in range (N_Malhas): #while dos vizinhos
                Vizinho = S_ch.copy()
                direcao = random.randint(0, 1) #1 pra direita e 0 pra esquerda
                if direcao == 0:
                    direcao = -1*i
                else:
                    direcao = 1*i
                
                #print (NomesChaves_aux)
                index = NomesChaves_aux.index(Vizinho[v])
                novoindex = index + direcao
        
                if (novoindex > N_Chaves_Malha_index[v]):
                    Vizinho[v] = NomesChaves_aux[novoindex - N_Chaves_Malha[v]]
                elif (novoindex < (N_Chaves_Malha_index[v]-N_Chaves_Malha[v])):
                    Vizinho[v] = NomesChaves_aux[N_Chaves_Malha_index[v]]
                else:
                    Vizinho[v] = NomesChaves_aux[novoindex]
                        
                fit = FitnessIndividuo_ch(Vizinho)
                if Radialidade():
                    Vizinhos.append(Vizinho)
                    FitVizinhos.append(fit)
                Vizinho = []
                        
        return Vizinhos, FitVizinhos
    
    #esta função faz o teste do S_ch que o programa entregou para verificar se foi a solução ótima para aquele sistema    
    def TestaSolucaoOtima (S_ch, N_Barras_Sistema):
        if N_Barras_Sistema == 33:
            S_otima = ['S7', 'S37', 'S9', 'S32', 'S14']
        elif N_Barras_Sistema == 69:
            S_otima = ["S69", "S14", "S55", "S61", "S70"]
        elif N_Barras_Sistema == 94:
            S_otima = ["S55","S7","S86","S72","S13","S89","S90","S83","S92","S39","S34","S42","S62"]
        
        ChavesComuns = list(set(S_otima) & set(S_ch))
        
        if len (ChavesComuns) == N_Malhas:
            return 1 #solução ótima
        else:
            return 0
    
    #verifica se as tensões das barras estão nos limites especificados
    #e envia também o valor da menor tensão e da maior tensão
    def Limites_tensao (tensao_especificada):
        limite_superior = tensao_especificada*1.05
        limite_inferior = tensao_especificada*0.93
        
        tensoes = objeto.get_tensoes_DSS()
        min_tensao_pu = min(tensoes)/tensao_especificada
        max_tensao_pu = max(tensoes)/tensao_especificada
        if (min(tensoes) < limite_inferior) or (max(tensoes) > limite_superior):
            return 0, min_tensao_pu, max_tensao_pu #se quebrar os limites, retorna 0
            print ("min_tensao_pu", min_tensao_pu)
            print (tensoes)
        
        return 1, min_tensao_pu, max_tensao_pu #se nao quebrar os limites, retorna 1 

    #verifica se a configuração não possui malhas e se está atendendo a todas as cargas
    def Radialidade ():
        if (objeto.num_isolated_loads_DSS() == 0) and (objeto.num_loops_DSS() == 0):
            return 1 #radial
        else:
            return 0 #nao radial
                
                
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
                self.dssTopology = self.dssCircuit.Topology
                self.dssSwtControl = self.dssCircuit.SwtControls
        
        def n_loops_DSS(self):
            return self.dssTopology.NumLoops
    
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
        
        def get_tensoes_DSS(self):
            return self.dssCircuit.AllBusVmag
    
        def num_isolated_loads_DSS(self): #numero de cargas isoladas
            return self.dssTopology.NumIsolatedLoads
    
        def num_loops_DSS(self): #numero de malhas no circuito
            return self.dssTopology.NumLoops
        
        def ativa_elemento(self, nome_elemento):

            # Ativa elemento pelo seu nome completo Tipo.Nome
            self.dssCircuit.SetActiveElement(nome_elemento)
            
            # Retonar o nome do elemento ativado
            #return self.dssCktElement.Name
            
        def ativa_barra (self, nome_barra):
            self.dssCircuit.SetActiveBus (nome_barra)
        
        def get_potencias_elemento(self):
            return self.dssCktElement.Powers
    
        def get_tensoes_elemento(self):
            return self.dssCktElement.VoltagesMagAng
        
        def get_barras_elemento(self):
    
            barras = self.dssCktElement.BusNames
    
            barra1 = barras[0]
            barra2 = barras[1]
    
            return barra1, barra2
        
        def get_names_lines(self):
            return self.dssLines.AllNames
        
        def get_buildYmatrix(self):
            NumNodes = self.dssCircuit.NumNodes
            Y = np.asarray(self.dssCircuit.SystemY).view(dtype=np.complex).reshape((NumNodes, NumNodes))
            return Y
        
        def get_dados_barras (self):
        
            AllBusNames = self.dssCircuit.AllBusNames
            #print (AllBusNames)
            carga_kW = []
            carga_kvar = []
            load_list = []
            tensoes_barras = []
            for i in range (len(AllBusNames)):
                self.dssCircuit.SetActiveBus (AllBusNames[i])
                load_list.append (self.dssBus.LoadList)
                tensoes_barras.append (self.dssBus.VMagAngle)
                if (load_list[i][0]):
                    self.dssCircuit.SetActiveElement(load_list[i][0])
                    carga_kW.append (self.dssCktElement.Powers[0])
                    carga_kW.append (self.dssCktElement.Powers[2])
                    carga_kW.append (self.dssCktElement.Powers[4])
                    carga_kvar.append (self.dssCktElement.Powers[1])
                    carga_kvar.append (self.dssCktElement.Powers[3])
                    carga_kvar.append (self.dssCktElement.Powers[5])
                else:
                    carga_kW.append (0)
                    carga_kW.append (0)
                    carga_kW.append (0)
                    carga_kvar.append (0)
                    carga_kvar.append (0)
                    carga_kvar.append (0)
                
            #print (AllBusNames)
            #print (load_list)
            #print (carga_kW)
            #print (carga_kvar)
            #print (tensoes_barras)
            
            
            return tensoes_barras, carga_kW, carga_kvar, load_list
                 
    
    if __name__ == "__main__":
        print ("""Autor: Luiz Otávio""")
    
        # Criar um objeto da classe DSS
        # Inserir aqui o nome da pasta onde encontra-se o seu .dss
        # Caso haja espaços no endereço, colocar entre []
        objeto = DSS(address_opendss)
        objeto.compile_DSS()
        print (u"Versão do OpenDSS: " + objeto.versao_DSS() + "\n")
                
        NomesTodasChaves = Ajuste_SistemaNBarras (N_Barras_Sistema) #obtém o vetor com todas as chaves de todas as malhas (incluindo as chaves repetidas)
        
        S_ch = S_inicial_PFO(NomesTodasChaves)
        """if (N_Barras_Sistema == 33):
            S_ch = ['S7','S28','S11','S17','S14']
            print ("Iniciando com a solução inicial Distância Elétrica pronta!")
        elif (N_Barras_Sistema == 69):
            S_ch = ['S58','S42','S14','S20','S26']
            print ("Iniciando com a solução inicial Distância Elétrica pronta!")
        elif (N_Barras_Sistema == 94):
            S_ch = ["S96","S88","S91","S87","S89","S94","S95","S93","S92","S90","S84","S85","S86"]"""

        NomesChaves, N_Chaves_Malha, N_Malhas = OrganizaChaves_com_S_ch (NomesTodasChaves, S_ch)
        FitS = FitnessIndividuo_ch(S_ch)
        print ("Configuração inicial: ", S_ch)
        
        BestFit = FitS
        BestS_ch = S_ch
        Iter = 0
        
        status_tensao, min_tensao, max_tensao = Limites_tensao(Tensao_Sistema(N_Barras_Sistema))
        if (status_tensao == 1):
            print ("Tensao nos limites")
            tensao = 1
        else:
            print ("Tensao fora dos limites")
            tensao = 0
        if (Radialidade()):
            print("Sistema radial")
            radialidade = 1
        else:
            print ("Sistema nao radial")
            print ("num loops", objeto.num_loops_DSS())
            radialidade = 0
    
        print ("Fit inicial: ", FitS)
        
        i = 0
        BestIter = 1        
        
        lista_tabu = [-1]*sum(N_Chaves_Malha)
        novo_tabu = [-1]*sum(N_Chaves_Malha)
        lista_tabu_completa = []
                                            
        while (i-BestIter) <= BTMax:
            
            i+=1
            #print ("================NOVA ITERAÇÃO=================", i)
            #print ("Partindo com Solução Corrente:", S_ch)
            
            VizS_ch, FitVizS = Vizinhos_LR_N (S_ch, N_Malhas, N_Chaves_Malha, N_Barras_Sistema)
            
            lista_tabu = novo_tabu    
            
            #vai adicionando os movimentos de cada S selecionado à lista_tabu_completa;
            #conforme ela vai preenchendo os movimentos mais antigos vão sendo eliminados
            if (len(lista_tabu_completa)==T): 
                lista_tabu_completa.pop(0)
                lista_tabu_completa.append(lista_tabu)
            else:
                lista_tabu_completa.append(lista_tabu)
                
            #organiza o fit, os movimentos e os vizinhos de acordo com o fit (do melhor pro pior)    
            fitorg, vizorg_ch = SortFitVizS_VizS_ListaMovimentos(FitVizS, N_Malhas, VizS_ch)
                        
            #for j in range(len(vizorg_ch)):
            #   print ("Vizinho", j, VizS_ch[j], "Fit:", FitVizS[j])
        
            #PARTE DA ESCOLHA DO NOVO S:
            if (fitorg[0] < BestFit): #função de aspiração: caso o melhor vizinho seja a melhor solução já encontrada mas estiver na LT, aceita mesmo assim 
                S_ch = vizorg_ch[0]
                fit_novo_S = fitorg[0]
                novo_tabu = S_ch
            else: #caso esteja na lista tabu, entra pro processo de escolha
                lista_tabu_completa_flat = [item for sublist in lista_tabu_completa for item in sublist]
                for m in range (len(vizorg_ch)):
                    if (any(n in lista_tabu_completa_flat for n in vizorg_ch[m]) == False):
                        #print ("O vizinho escolhido foi o", m)
                        S_ch = vizorg_ch[m]
                        fit_novo_S = fitorg[m]
                        novo_tabu = S_ch
                        break
            
            #se o fit do novo S for o melhor já encontrado:
            if (fit_novo_S < BestFit):
                BestS_ch = S_ch
                BestIter = i
                #print ("Novos BestFit e BestS encontrados")
                BestFit =  FitnessIndividuo_ch(BestS_ch)
                status_tensao, min_tensao, max_tensao = Limites_tensao(Tensao_Sistema(N_Barras_Sistema))
                if (status_tensao == 1):
                    #print ("Tensao nos limites")
                    tensao = 1
                else:
                    #print ("Tensao fora dos limites")
                    tensao = 0
                if (Radialidade()):
                    #print("Sistema radial")
                    radialidade = 1
                else:
                    #print ("Sistema nao radial")
                    radialidade = 0
                
            if (i == MaxIter):
                break
            #print ("Novo S:", S_ch, "Fit:", fit_novo_S)
            #print ("Lista Tabu: ", novo_tabu)
            
            #print ("BestS: ", BestS_ch, "BestFit:", BestFit)
            
            #reorganiza as chaves para dar uma aleatoriedade ao sistema;
            #além disso, evita que as chaves fiquem engessadas em uma malha só sempre
            NomesChaves, N_Chaves_Malha, N_Malhas = OrganizaChaves_com_S_ch (NomesTodasChaves, S_ch)
            
            VizS_ch.clear()
            FitVizS.clear()
            fitorg.clear()
            FitVizS.clear()
            vizorg_ch.clear()
        
        
        print ("--------Resultado da rodada", k, "--------")
        print ("BestIter: ", BestIter)
        print ("Melhor configuração: ", BestS_ch, "Fit: ", BestFit)
        if (tensao==1):
            print ("Tensão: OK")
        else:
            print ("Tensão: erro")
        if (radialidade==1):
            print ("Radialidade: OK")
        else:
            print ("Radialidade: erro")
            
        if (tensao == 1 and radialidade == 1):
            k+=1
            MelhoresFitsGerais.append(BestFit)
            MelhoresSGerais.append(BestS_ch)
            MinimasTensoesGerais.append(min_tensao)
            MaximasTensoesGerais.append(max_tensao)
            solucao_otima = TestaSolucaoOtima(BestS_ch, N_Barras_Sistema)
            if solucao_otima == 1:
                solucoesotimas += 1
                print ("Solução ótima encontrada!")
            end = time.time()
            TempoGasto.append(end-start)
            
        NomesChaves = []
        NomesTodasChaves = []
        N_Chaves_Malha = []
        BestS_ch = []
        lista_tabu_completa_flat = []
        lista_tabu_completa = []

print ("================RESULTADO FINAL================")
for i in range (len(MelhoresFitsGerais)):
    print ("Config: ", MelhoresSGerais[i])
    print ("Fit: ", round (MelhoresFitsGerais[i],3), "kW")
    print ("Tensao minima: ", round (MinimasTensoesGerais[i],4), "pu")           
    print ("Tensão máxima: ", round (MaximasTensoesGerais[i],4), "pu")
    print ("Tempo gasto: ", TempoGasto[i])
    
media_perdas = sum(MelhoresFitsGerais)/len(MelhoresFitsGerais)
desvio_padrao = np.std(MelhoresFitsGerais) 

print ("Média de perdas para",N_Vezes,"configurações obtidas para sistema de", 
       N_Barras_Sistema, "barras: ", round(media_perdas,3))

print ("O sistema obteve", solucoesotimas,"soluções ótimas para", len(MelhoresFitsGerais), "soluções entregues")   

print ("Tempo gasto total: ", sum(TempoGasto))

valores = [N_Vezes, N_Barras_Sistema, BTMax, T, MaxIter, solucoesotimas, media_perdas, desvio_padrao,
           sum(TempoGasto), sum(TempoGasto)/len(TempoGasto), versao_codigo]

ws.append(valores)
    
wb.save(r"C:\Users\luiz3\Google Drive (luiz.filho@cear.ufpb.br)\Tese\Resultados_TS.xlsx")

