MelhoresFitsGerais = []
MelhoresSGerais = []

for i in range (1):
    
    import numpy as np
    import win32com.client
    from datetime import datetime
    import random 
    
    
    #esta função organiza as chaves de cada malha sorteando aleatoriamente para cada malha
    #cada uma das chaves empatadas (que per)
    def OrganizaChaves_33bar ():
        
        ChavesMalha = [0,0,0,0,0]
        N_Chaves_Malha = []
        
        ChavesMalha[0] = ["S2", "S3","S4","S5","S6","S7","S33","S20","S19","S18"]
        ChavesMalha[1] = ["S22","S23","S24","S37","S28","S27","S26","S25","S5","S4","S3"]
        ChavesMalha[2] = ["S33","S8","S9","S10","S11","S35","S21"]
        ChavesMalha[3] = ["S25","S26","S27","S28","S29","S30","S31","S32","S36","S17","S16","S15","S34","S8","S7","S6"]
        ChavesMalha[4] = ["S34","S14","S13","S12","S11","S10","S9"]
        
        for i in range (len(ChavesMalha)):
            for j in range (len(ChavesMalha)):
                if (i!=j):
                    ChavesComuns = list(set(ChavesMalha[i]) & set(ChavesMalha[j]))
                    if ChavesComuns:
                        for k in range (len(ChavesComuns)):
                            Aleatorio = np.random.randint(2) 
                            if (Aleatorio == 0): #se 0 a chave fica na malha i, se 1 fica na malha j
                                ChavesMalha[j].remove(ChavesComuns[k])
                            elif (Aleatorio == 1):
                                ChavesMalha[i].remove(ChavesComuns[k])
                        ChavesComuns.clear()
        
        for i in range (len(ChavesMalha)):
            N_Chaves_Malha.append(len(ChavesMalha[i]))
        
        N_Malhas = len(ChavesMalha)
        
        NomesChaves = sum(ChavesMalha, [])
        
        
        return NomesChaves, N_Chaves_Malha, N_Malhas
        
    #esta função organiza o Fit em ordem crescente e consequentemente os Vizinhos e a lista de movimentos 
    #em ordem crescente de acordo com o Fit já ordenado 
    def SortFitVizS_VizS_ListaMovimentos(FitVizS, N_Malhas, VizS, lista_movimentos_separados):
        
        Fit_Sorted = sorted(FitVizS, reverse=False)
        
        VizS_Sorted = []
        movimentos_Sorted = []
        
        for i in range (len(FitVizS)):
            index = FitVizS.index(Fit_Sorted[i])
            VizS_Sorted.append(VizS[index])
            movimentos_Sorted.append(lista_movimentos_separados[index])
            
        
        return Fit_Sorted, VizS_Sorted, movimentos_Sorted
            
    #esta função cria um novo S e uma nova lista tabu. ela analisa malha a malha se 
    #em nenhum momento o vizinho de melhor fit é tabu com o tabu do S anterior.
    #caso uma malha seja tabu, a função pega o movimento do próximo vizinho de melhor fit (segundo, terceiro etc)
    #de forma a criar um NovoS totalmente renovado em relação ao anterior.
    def NovoS (vizinhos, N_Chaves_Malha, lista_tabu_completa, movimentos_Sorted):
        
                 
        lista_tabu_completa_transposta = np.transpose(lista_tabu_completa)
        novo_tabu = []
        novo_S = []
        for i in range (len(movimentos_Sorted)):
            for j in range (len(movimentos_Sorted[i])):
                if (movimentos_Sorted[j][i] not in lista_tabu_completa_transposta[j]):
                    novo_tabu.append(movimentos_Sorted[j][i])
                    break
                    
        novo_S = [1]*sum(N_Chaves_Malha)
        for i in novo_tabu:
            novo_S[i] = 0
        
        fit = FitnessIndividuo (novo_S, NomesChaves)

        
        return novo_S, novo_tabu, fit
          
    
    def DecodificadorChaves (Individuo, NomesChaves):
        ChavesAbertas = []
        for i in range(len(NomesChaves)):
            if (Individuo[i]==0):
                ChavesAbertas.append(NomesChaves[i])
    
        return ChavesAbertas

    def CodificadorChaves(ChavesAbertas, NomesChaves):
        
        Individuo = [1]*len(NomesChaves)
        for i in range(len(ChavesAbertas)):
            index = NomesChaves.index(ChavesAbertas[i])
            Individuo[index] = 0
        
        return Individuo
    
    
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
        FitVizinhos3 = []
    
        #atualmente estou escolhendo um aleatorio em cada malha. no primeiro vizinho esse aleatorio vai pra 0 
        #e onde tá 0 vai pra 1
        #nos demais vizinhos vai andando com o 0 pra frente de onde estava no primeiro vizinho
        i = 0
        while i < N_Malhas: #for dos vizinhos
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
            fit = FitnessIndividuo(Vizinho, NomesChaves)
            if (Radialidade()):
                Vizinhos2.append(np.array(Vizinho))
                Vizinhos.append(Vizinho)
                FitVizinhos3.append(fit)
                Vizinhos3.append(Vizinhos2[i].tolist())
                i+=1
            else:
                changes_list = changes_list[:len(changes_list)-N_Malhas]
                 
            Vizinhos2.append(np.array(Vizinho))
            Vizinhos.append(Vizinho)
        
            
        for i in range(len(Vizinhos2)):
            Vizinhos3.append(Vizinhos2[i].tolist())
        
        changes_list_separated = np.array_split(changes_list, N_Malhas)
                                          
        return Vizinhos3, FitVizinhos3, changes_list_separated
                    
    def GeraS(N_Chaves_Malha, N_Malhas): #gera um S aleatório
        
        S = [1]*sum(N_Chaves_Malha)
        
        inicio = 0
        for i in range(N_Malhas):
            random.seed(datetime.now())
            x = random.randint(inicio, inicio + N_Chaves_Malha[i]-1)
            S[x] = 0
            inicio = inicio + N_Chaves_Malha[i]
        
        return S
    
    def Limites_tensao (tensao_especificada):
        limite_superior = tensao_especificada*1.05
        limite_inferior = tensao_especificada*0.93
        
        tensoes = objeto.get_tensoes_DSS()
        if (min(tensoes) < limite_inferior) or (max(tensoes) > limite_superior):
            return 0 #se quebrar os limites, retorna 0
        
        return 1 #se nao quebrar os limites, retorna 1 

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
        N_Chaves_Malha.append(8) #numero de chaves da malha 4
        N_Chaves_Malha.append(5) #numero de chaves da malha 5
        
        S = [1,1,1,1,1,1,0,1,1,1,
             1,1,1,0,1,1,1,1,
             1,1,1,0,1,
             1,1,1,1,0,1,1,1,
             0,1,1,1,1]
        
        
        """NomesChaves, N_Chaves_Malha, N_Malhas = OrganizaChaves_33bar()
        
        S = GeraS(N_Chaves_Malha, N_Malhas)"""
        
       
        print ("N_Chaves_Malha", N_Chaves_Malha)
    
        FitS = FitnessIndividuo(S, NomesChaves)
        print ("Configuração inicial: ", DecodificadorChaves(S, NomesChaves))
        
        BestFit = FitS
        BestS = S
        Iter = 0
        BTMax = 30
    
        print ("Fit inicial: ", FitS)
        
        i = 0
        BestIter = 1
        T = 4 #tamanho da lista tabu
        
        FitVizS = []
        lista_tabu = [-1]*sum(N_Chaves_Malha)
        novo_tabu = [-1]*sum(N_Chaves_Malha)
        ChavesAbertasVizS = []
        lista_tabu_completa = []
        
        melhorvizinho = 0
        vizinhogerado = 0
        TodosS = []
        
        while (i-BestIter) <= BTMax:
            

            i+=1
            print ("================NOVA ITERAÇÃO=================", i)
            print ("Partindo com Solução Corrente:", DecodificadorChaves(S, NomesChaves))
            VizS, FitVizS, lista_movimentos_separados = Vizinhos(S, N_Malhas, N_Chaves_Malha)
            #lista_movimentos é um array que guarda todos os movimentos feitos em cada malha em cada vizinho
            #ou seja, lista_movimentos = [0, 1, 2, 4, 5, 6], por exemplo, 0 1 2 foram os movimentos para as malhas 1, 2 e 3 do primeiro vizinhi
            #enquanto 4, 5 6 foram movimentos das malhas 1 2 e 3 para o segundo vizinho etc
            #lista_movimentos_separados gera os movimentos de cada vizinho em listas diferentes, tipo [[0, 1, 2],[3, 4, 5]]
            
            
            lista_tabu = novo_tabu    
            
            if (len(lista_tabu_completa)==T):
                lista_tabu_completa.pop(0)
                lista_tabu_completa.append(lista_tabu)
            else:
                lista_tabu_completa.append(lista_tabu)

                
            fitorg, vizorg, moviorg = SortFitVizS_VizS_ListaMovimentos(FitVizS, N_Malhas, VizS, lista_movimentos_separados)
                
            if (i>1):
                S_gerado, movi_S_gerado, fit_S_gerado = NovoS(moviorg, N_Chaves_Malha, lista_tabu_completa, moviorg)
                VizS.append(S_gerado)
                lista_movimentos_separados.append(movi_S_gerado)
                FitVizS.append(fit_S_gerado)
                
            fitorg, vizorg, moviorg = SortFitVizS_VizS_ListaMovimentos(FitVizS, N_Malhas, VizS, lista_movimentos_separados)
            
            for j in range(len(vizorg)):
                #print ("Vizinho: " ,VizS[j])
                print ("Vizinho", j, DecodificadorChaves(vizorg[j], NomesChaves), "Fit:", fitorg[j])
        
            S = vizorg[0]
            fit_novo_S = fitorg[0]
            novo_tabu = np.array(moviorg[0]).tolist()
            
            TodosS.append(S)
            
            if (BestFit > fit_novo_S):
                BestFit =  fit_novo_S
                BestS = DecodificadorChaves(S, NomesChaves)
                BestS2 = S
                BestIter = i
                print ("Novos BestFit e BestS encontrados")
                if (Limites_tensao(7309)):
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
              
                
            
            print ("Novo S:", DecodificadorChaves(S, NomesChaves), "Fit:", fit_novo_S)
            #print ("Lista Tabu: ", novo_tabu)
            
            print ("BestS: ", BestS, "BestFit:", BestFit)
            
            
            VizS.clear()
            FitVizS.clear()
            lista_movimentos_separados.clear()
            fitorg.clear()
            FitVizS.clear()
            VizS.clear()
            #fitvizorg.clear()
            vizorg.clear()
            moviorg.clear()
            
        
        MelhoresFitsGerais.append(BestFit)
        MelhoresSGerais.append(BestS)
        print ("Resultado desta rodada======== ")
        print ("BestIter: ", BestIter)
        print ("Melhor configuração: ", BestS, "Fit: ", BestFit)
        if (tensao==1):
            print ("Tensão: OK")
        else:
            print ("Tensão: erro")
        if (radialidade==1):
            print ("Radialidade: OK")
        else:
            print ("Radialidade: erro")
        #for i in range(len(S)):
         #   objeto.switch_elemento(NomesChaves[i], S[i])
            
        #objeto.solve_DSS_snapshot() 
        
        #for i in range(len(MelhoresS)):
    
    #print ("numero de novo_S do melhor vizinho: ", melhorvizinho)
    # ("numero de novo_S gerado: ", vizinhogerado)
    
print ("================RESULTADO FINAL================")
for i in range (len(MelhoresFitsGerais)):
    print ("Config: ", MelhoresSGerais[i], 
           "Fit: ", MelhoresFitsGerais[i])
    
"""print ("Melhor configuração: ", MelhoresFitsGerais)
print ("Média de perdas: ", sum(MelhoresFitsGerais)/len(MelhoresFitsGerais))
moda = statistics.mode(MelhoresFitsGerais)
print ("Moda: ", moda)"""
        
            
    
    
    
