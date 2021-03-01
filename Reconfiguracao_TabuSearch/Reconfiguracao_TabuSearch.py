"""Especifique quantas vezes você quer que o programa rode:"""
N_Vezes = 5
"""Especifique o endereço do arquivo Opendss (dentro dos colchetes):"""
address_opendss = r"[C:\Users\luiz3\Google Drive (luiz.filho@cear.ufpb.br)\Tese\Código\sistemas\69_barras.dss]"

"""Especifique o número de barras do sistema que você quer trabalhar (33 ou 69):"""
N_Barras_Sistema = 69

"""Especifique quantas iterações sem melhora o sistema deve fazer no máximo"""
BTMax = 100

MelhoresFitsGerais = []
MelhoresSGerais = []
MinimasTensoesGerais = []
MaximasTensoesGerais = []

k = 0
for i in range (N_Vezes):
#while k < N_Vezes:
    
    import numpy as np
    import win32com.client
    from datetime import datetime
    import random 
    
    def Ajuste_SistemaNBarras (N_Barras):
        
        if (N_Barras == 69):
            ChavesMalha = [0,0,0,0,0]
            ChavesMalha[0] = ["S4","S5","S6","S7","S8","S46","S47","S48","S49","S52","S53",
                          "S54","S55","S56","S57","S58","S72"]
            ChavesMalha[1] = ["S3","S4","S5","S6","S7","S8","S9","S10","S35","S36","S37",
                          "S38","S39","S40","S41","S42","S69"]
            ChavesMalha[2] = ["S11","S12","S13","S14","S43","S44","S45","S69","S71"]
            ChavesMalha[3] = ["S13","S14","S15","S16","S17","S18","S19","S20","S70"]
            ChavesMalha[4] = ["S9","S10","S11","S12","S21","S22","S23","S24","S25","S26","S52","S53",
                      "S54","S55","S56","S57","S58","S59","S60","S61","S62","S63",
                      "S64","S73"]
            return ChavesMalha
        elif (N_Barras == 33):
            ChavesMalha = [0,0,0,0,0]
            ChavesMalha[0] = ["S2", "S3","S4","S5","S6","S7","S33","S20","S19","S18"]
            ChavesMalha[1] = ["S22","S23","S24","S37","S28","S27","S26","S25","S5","S4","S3"]
            ChavesMalha[2] = ["S33","S8","S9","S10","S11","S35","S21"]
            ChavesMalha[3] = ["S25","S26","S27","S28","S29","S30","S31","S32","S36","S17","S16","S15","S34","S8","S7","S6"]
            ChavesMalha[4] = ["S34","S14","S13","S12","S11","S10","S9"]
            return ChavesMalha
            
    
    #esta função roda o sistema com todas as chaves fechadas e então seleciona as chaves com
    #menor fluxo de potência em cada malha; a função observa tambem caso uma chave seja
    #a chave de menor fluxo de duas malhas; caso afirmativo, a chave é selecionada para uma
    #malha e na outra malha é selecionada a chave que tem menor fluxo que não seja em comum com
    #a primeira malha
    def S_inicial(NomesChaves):
        
        objeto.solve_DSS_snapshot()       
        potencias_ativas_trifasicas = []
        potencias_malha_aux = []
        
        #essa parte obtém o fluxo de cada uma das linhas e armazena no vetor potencias_ativas_trifasicas
        #é válido comentar que esse vetor corresponde com o NomesChaves
        for i in range (len(NomesChaves)): #for de cada uma das malhas
            potencias_malha_aux = []
            for j in range (len(NomesChaves[i])): #for das chaves dentro de cada malha
                num_linha = int(NomesChaves[i][j][1:]) #pega o numero da linha a que se refere
                objeto.ativa_elemento("Line.L"+str(num_linha)) #ativa a determinada linha
                potencias = (objeto.get_potencias_elemento()) #pega todas as potencias da linha
                potencias_malha_aux.append(abs(potencias[0] + potencias[2] + potencias[4])) #soma somente as potências ativas entrando na linha
            potencias_ativas_trifasicas.append(potencias_malha_aux)
        
        
        
        S = []
        for i in range (len(potencias_ativas_trifasicas)): #for de cada uma das malhas
            potencias_sorted = sorted(potencias_ativas_trifasicas[i], reverse=False) #ordena os fluxos de potência de forma crescente para uma determinada malha
            menor_fluxo_chave_malha = potencias_sorted[0]
            index_menor_fluxo_chave_malha = potencias_ativas_trifasicas[i].index(menor_fluxo_chave_malha) #obtém o index do menor fluxo no vetor original potencias_ativas_trifasicas
            chave = NomesChaves[i][index_menor_fluxo_chave_malha] #obtém a chave no vetor NomesChaves
            
            if (chave in S): #se a chave já tiver sido escolhida 
                index_malhacomum = S.index(chave) #descobre a malha que a chave selecionada está em comum
                ChavesComuns = [value for value in NomesChaves[i] if value in NomesChaves[index_malhacomum]] #obtém as chaves em comum entre a malha I e a malha da chave em comum
                ChavesNaoComunsMalhai = [x for x in NomesChaves[i] if x not in ChavesComuns] #obtém as chaves que não são comuns entre a malha I e a malha da chave em comum
                potencias_aux = []
                for j in range (len(ChavesNaoComunsMalhai)): 
                    index_aux = NomesChaves[i].index(ChavesNaoComunsMalhai[j]) #encontra o index das chaves não comuns dentro do vetor NomesChaves
                    potencias_aux.append(potencias_ativas_trifasicas[i][index_aux]) #adiciona ao vetor potencias_aux os fluxos das chaves não comuns
                    potencias_aux_sorted = sorted(potencias_aux, reverse=False) #ordena as potencias em ordem crescente
                    
                menor_fluxo_chave_malha = potencias_aux_sorted[0] #pega o menor fluxo entre as chaves não comuns
                index_menor_fluxo_chave_malha = potencias_ativas_trifasicas[i].index(menor_fluxo_chave_malha)
                chave = NomesChaves[i][index_menor_fluxo_chave_malha] #obtém no vetor NomesChaves a chave de menor fluxo
            
            S.append(chave)
        
        print ("S inicial: ", S)        
        
        return S
    
    
    #esta função recebe o vetor ChavesMalha que contém todas as chaves de todas as malhas (logo, chaves repetidas)
    #e organiza de forma que cada malha tenha apenas suas chaves exclusivas e realiza um sorteio
    #para que cada chave que pertence a duas malhas seja designada a apenas uma malha
    #logo, o vetor NomesChaves possui os nomes de todas as chaves em ordem de malha
    #O N_Chaves_Malha diz a quantidade de chaves da malha 1, 2, 3 etc.
    #n N_Malhas é somente o número de malhas do sistema
    def OrganizaChaves (ChavesMalha):
        
        N_Chaves_Malha = []
        
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
    def NovoS (vizinhos, N_Chaves_Malha, lista_tabu_completa, movimentos_Sorted, fit_vizinhos):
        
        lista_tabu_completa_joined = [j for i in lista_tabu_completa for j in i]
        novo_tabu = []
        novo_S = []
        for i in range (len(movimentos_Sorted)):
            for j in range (len(movimentos_Sorted[i])):
                if (lista_tabu_completa_joined.count(movimentos_Sorted[j][i])==0):
                    novo_tabu.append(movimentos_Sorted[j][i])
                    break
                    
        novo_S = [1]*sum(N_Chaves_Malha)
        for i in novo_tabu:
            novo_S[i] = 0
        
        fit = FitnessIndividuo (novo_S, NomesChaves)
        if (Radialidade()):
            return novo_S, novo_tabu, fit
        else:
            return vizinhos[0], movimentos_Sorted[0], fit_vizinhos[0]
            
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
    
    #calcula o fitness de um S binário qualquer.
    def FitnessIndividuo(S, NomesChaves):
        
        for i in range(len(S)):
            objeto.switch_elemento(NomesChaves[i], S[i])
            
        objeto.solve_DSS_snapshot()       
        return objeto.get_line_losses()
    

    #nesta função são gerados o número de vizinhos igual ao número de malhas do sistema
    #são gerados N vizinhos que cumpram o requisito de radialidade do sistema
    #todos os vizinhos são aleatórios; ou seja, para um dado vizinho, a chave aberta em 
    #cada uma das malhas é aleatória.
    #como saída, a função envia os Vizinhos na forma binária, os fits desses vizinhos
    #e as posições das chaves que foram modificadas (changes_list_separated)
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
        while i < N_Malhas: #while dos vizinhos
            inicio = 0
            Vizinho = [1]*len(S)
            for j in range (N_Malhas): #for das malhas  
                x = random.randint(inicio, inicio + N_Chaves_Malha[j]-1) #escolhe aleatoriamente uma posição de bit dentro da solução S para que seja aberta aquela chave daquela posição
                if (Vizinho[x]==1):
                    Vizinho[x] = 0
                changes_list.append(x)
                inicio = inicio + N_Chaves_Malha[j] #atualiza o range de onde sortear a chave (por causa da malha)
                    
            fit = FitnessIndividuo(Vizinho, NomesChaves) #calcula o fit para que seja possível verificar a radialidade também
            if (Radialidade()):
                Vizinhos2.append(np.array(Vizinho))
                Vizinhos.append(Vizinho)
                FitVizinhos3.append(fit)
                Vizinhos3.append(Vizinhos2[i].tolist())
                i+=1
            else:
                changes_list = changes_list[:len(changes_list)-N_Malhas]
        
        changes_list_separated = np.array_split(changes_list, N_Malhas)
                                          
        return Vizinhos3, FitVizinhos3, changes_list_separated
     
    #gera um S aleatório 
    def GeraS(N_Chaves_Malha, N_Malhas): 
        
        S = [1]*sum(N_Chaves_Malha)
        
        inicio = 0
        for i in range(N_Malhas):
            random.seed(datetime.now())
            x = random.randint(inicio, inicio + N_Chaves_Malha[i]-1)
            S[x] = 0
            inicio = inicio + N_Chaves_Malha[i]
        
        return S
    
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
                 
    
    if __name__ == "__main__":
        print ("""Autor: Luiz Otávio""")
    
        # Criar um objeto da classe DSS
        # Inserir aqui o nome da pasta onde encontra-se o seu .dss
        # Caso haja espaços no endereço, colocar entre []
        objeto = DSS(address_opendss)
        objeto.compile_DSS()
        print (u"Versão do OpenDSS: " + objeto.versao_DSS() + "\n")
                
        NomesTodasChaves = Ajuste_SistemaNBarras (N_Barras_Sistema) #obtém o vetor com todas as chaves de todas as malhas (incluindo as chaves repetidas)
    
        S_start = S_inicial(NomesTodasChaves) #obtém o S inicial pelo método do menor fluxo nas malhas
    
        NomesChaves, N_Chaves_Malha, N_Malhas = OrganizaChaves(NomesTodasChaves) 
        
        S = CodificadorChaves(S_start, NomesChaves)
               
        FitS = FitnessIndividuo(S, NomesChaves)
        print ("Configuração inicial: ", DecodificadorChaves(S, NomesChaves))
        
        BestFit = FitS
        BestS = DecodificadorChaves(S, NomesChaves)
        Iter = 0
        
        status_tensao, min_tensao, max_tensao = Limites_tensao(7309)
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
            radialidade = 0
    
        print ("Fit inicial: ", FitS)
        
        i = 0
        BestIter = 1
        T = 3 #tamanho da lista tabu
        
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
            
            #vai adicionando os movimentos de cada S selecionado à lista_tabu_completa;
            #conforme ela vai preenchendo os movimentos mais antigos vão sendo eliminados
            if (len(lista_tabu_completa)==T): 
                lista_tabu_completa.pop(0)
                lista_tabu_completa.append(lista_tabu)
            else:
                lista_tabu_completa.append(lista_tabu)
                
            #organiza o fit, os movimentos e os vizinhos de acordo com o fit (do melhor pro pior)    
            fitorg, vizorg, moviorg = SortFitVizS_VizS_ListaMovimentos(FitVizS, N_Malhas, VizS, lista_movimentos_separados)
            
            #caso já tenha passado a primeira iteração já é possível calcular o S gerado (levando em consideração a LT)
            if (i>1):
                S_gerado, movi_S_gerado, fit_S_gerado = NovoS(vizorg, N_Chaves_Malha, lista_tabu_completa, moviorg, fitorg)
                VizS.append(S_gerado)
                lista_movimentos_separados.append(movi_S_gerado)
                FitVizS.append(fit_S_gerado)
                print ("S gerado: ", DecodificadorChaves(S_gerado, NomesChaves))
            
            #reorganiza novamente agora com o novo vizinho (S gerado)
            fitorg, vizorg, moviorg = SortFitVizS_VizS_ListaMovimentos(FitVizS, N_Malhas, VizS, lista_movimentos_separados)
            
            for j in range(len(vizorg)):
                print ("Vizinho", j, DecodificadorChaves(vizorg[j], NomesChaves), "Fit:", fitorg[j])
        
            #PARTE DA ESCOLHA DO NOVO S:
            if (fitorg[0] < BestFit): #função de aspiração: caso o melhor vizinho seja a melhor solução já encontrada mas estiver na LT, aceita mesmo assim 
                S = vizorg[0]
                fit_novo_S = fitorg[0]
                novo_tabu = np.array(moviorg[0]).tolist()
            else: #caso esteja na lista tabu, entra pro processo de escolha
                lista_tabu_completa_flat = [item for sublist in lista_tabu_completa for item in sublist]
                for m in range (len(vizorg)):
                    if (any(n in lista_tabu_completa_flat for n in moviorg[m]) == False):
                        print ("O vizinho escolhido foi o", m)
                        S = vizorg[m]
                        fit_novo_S = fitorg[m]
                        novo_tabu = np.array(moviorg[m]).tolist() 
                        break
            
            #se o fit do novo S for o melhor já encontrado:
            if (BestFit > fit_novo_S):
                BestFit =  fit_novo_S
                BestS = DecodificadorChaves(S, NomesChaves)
                BestS2 = S
                BestIter = i
                print ("Novos BestFit e BestS encontrados")
                status_tensao, min_tensao, max_tensao = Limites_tensao(7309)
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
                    radialidade = 0
                
            
            print ("Novo S:", DecodificadorChaves(S, NomesChaves), "Fit:", fit_novo_S)
            #print ("Lista Tabu: ", novo_tabu)
            
            print ("BestS: ", BestS, "BestFit:", BestFit)
            
            #reorganiza as chaves para dar uma aleatoriedade ao sistema;
            #além disso, evita que as chaves fiquem engessadas em uma malha só sempre
            NomesChaves, N_Chaves_Malha, N_Malhas = OrganizaChaves(NomesTodasChaves)
            
            VizS.clear()
            FitVizS.clear()
            lista_movimentos_separados.clear()
            fitorg.clear()
            FitVizS.clear()
            VizS.clear()
            #fitvizorg.clear()
            vizorg.clear()
            moviorg.clear()
            
        
        print ("--------Resultado desta rodada-------- ")
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
            
        if (tensao == 1 and radialidade == 1):
            k+=1
            MelhoresFitsGerais.append(BestFit)
            MelhoresSGerais.append(BestS)
            MinimasTensoesGerais.append(min_tensao)
            MaximasTensoesGerais.append(max_tensao)
            
            
            
print ("================RESULTADO FINAL================")
for i in range (len(MelhoresFitsGerais)):
    print ("Config: ", MelhoresSGerais[i])
    print ("Fit: ", round (MelhoresFitsGerais[i],3), "kW")
    print ("Tensao minima: ", round (MinimasTensoesGerais[i],3), "pu")           
    print ("Tensão máxima: ", round (MaximasTensoesGerais[i],3), "pu")
    
print ("Média de perdas para as configurações obtidas para sistema de ", 
       N_Barras_Sistema, "barras: ", round(sum(MelhoresFitsGerais)/len(MelhoresFitsGerais),3))
        
            
    
    
    


