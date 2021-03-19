"""Especifique quantas vezes você quer que o programa rode:"""
N_Vezes = 40
"""Especifique o endereço do arquivo Opendss (dentro dos colchetes):"""
address_opendss = r"[C:\Users\luiz3\Google Drive (luiz.filho@cear.ufpb.br)\Tese\Código\sistemas\94_nos.dss]"

"""Especifique o número de barras do sistema que você quer trabalhar (33, 69 ou 94):"""
N_Barras_Sistema = 94

"""Especifique quantas iterações sem melhora o sistema deve fazer no máximo"""
BTMax = 200

"""Especifique o tamanho da Lista Tabu"""
T = 3

"""Número máximo de iterações"""
MaxIter = 1000

MelhoresFitsGerais = []
MelhoresSGerais = []
MinimasTensoesGerais = []
MaximasTensoesGerais = []
TempoGasto = []

solucoesotimas = 0
k = 0
for i in range (N_Vezes):
#while k < N_Vezes:
    
    import numpy as np
    import win32com.client
    from datetime import datetime
    import random 
    import copy
    import time
    
    start = time.time()
    
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
            NomesTodasChaves[0] = ["S84","S54","S53","S96","S63","S62","S61","S85","S6",
                                   "S7","S4"]
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
            
    
    #esta função roda o sistema com todas as chaves fechadas e então seleciona as chaves com
    #menor fluxo de potência em cada malha; a função observa tambem caso uma chave seja
    #a chave de menor fluxo de duas malhas; caso afirmativo, a chave é selecionada para uma
    #malha e na outra malha é selecionada a chave que tem menor fluxo que não seja em comum com
    #a primeira malha
    def S_inicial(NomesTodasChaves): 
        
        objeto.solve_DSS_snapshot()       
        potencias_ativas_trifasicas = []
        potencias_malha_aux = []
        
        #essa parte obtém o fluxo de cada uma das linhas e armazena no vetor potencias_ativas_trifasicas
        #é válido comentar que esse vetor corresponde com o NomesChaves
        for i in range (len(NomesTodasChaves)): #for de cada uma das malhas
            potencias_malha_aux = []
            for j in range (len(NomesTodasChaves[i])): #for das chaves dentro de cada malha
                num_linha = int(NomesTodasChaves[i][j][1:]) #pega o numero da linha a que se refere
                objeto.ativa_elemento("Line.L"+str(num_linha)) #ativa a determinada linha
                potencias = (objeto.get_potencias_elemento()) #pega todas as potencias da linha
                potencias_malha_aux.append(abs(potencias[0] + potencias[2] + potencias[4])) #soma somente as potências ativas entrando na linha
            potencias_ativas_trifasicas.append(potencias_malha_aux)
        
        S = []
        for i in range (len(potencias_ativas_trifasicas)): #for de cada uma das malhas
            potencias_sorted = sorted(potencias_ativas_trifasicas[i], reverse=False) #ordena os fluxos de potência de forma crescente para uma determinada malha
            menor_fluxo_chave_malha = potencias_sorted[0]
            index_menor_fluxo_chave_malha = potencias_ativas_trifasicas[i].index(menor_fluxo_chave_malha) #obtém o index do menor fluxo no vetor original potencias_ativas_trifasicas
            chave = NomesTodasChaves[i][index_menor_fluxo_chave_malha] #obtém a chave no vetor NomesTodasChaves
            
            if (chave in S): #se a chave já tiver sido escolhida 
                index_malhacomum = S.index(chave) #descobre a malha que a chave selecionada está em comum
                ChavesComuns = [value for value in NomesTodasChaves[i] if value in NomesTodasChaves[index_malhacomum]] #obtém as chaves em comum entre a malha I e a malha da chave em comum
                ChavesNaoComunsMalhai = [x for x in NomesTodasChaves[i] if x not in ChavesComuns] #obtém as chaves que não são comuns entre a malha I e a malha da chave em comum
                potencias_aux = []
                for j in range (len(ChavesNaoComunsMalhai)): 
                    index_aux = NomesTodasChaves[i].index(ChavesNaoComunsMalhai[j]) #encontra o index das chaves não comuns dentro do vetor NomesChaves
                    potencias_aux.append(potencias_ativas_trifasicas[i][index_aux]) #adiciona ao vetor potencias_aux os fluxos das chaves não comuns
                    potencias_aux_sorted = sorted(potencias_aux, reverse=False) #ordena as potencias em ordem crescente
                    
                menor_fluxo_chave_malha = potencias_aux_sorted[0] #pega o menor fluxo entre as chaves não comuns
                index_menor_fluxo_chave_malha = potencias_ativas_trifasicas[i].index(menor_fluxo_chave_malha)
                chave = NomesTodasChaves[i][index_menor_fluxo_chave_malha] #obtém no vetor NomesTodasChaves a chave de menor fluxo
            
            S.append(chave)
        
        print ("S inicial: ", S)        
        
        return S
    
    
    #esta função recebe o vetor ChavesMalha que contém todas as chaves de todas as malhas (logo, chaves repetidas)
    #e organiza de forma que cada malha tenha apenas suas chaves exclusivas e realiza um sorteio
    #para que cada chave que pertence a duas malhas seja designada a apenas uma malha
    #logo, o vetor NomesChaves possui os nomes de todas as chaves em ordem de malha
    #O N_Chaves_Malha diz a quantidade de chaves da malha 1, 2, 3 etc.
    #n N_Malhas é somente o número de malhas do sistema
    def OrganizaChaves (NomesTodasChaves):
    
        N_Chaves_Malha = []
        
        NomesTodasChaves_aux = copy.deepcopy(NomesTodasChaves)
        
        for i in range (len(NomesTodasChaves_aux)):
            for j in range (len(NomesTodasChaves_aux)):
                if (i!=j):
                    ChavesComuns = list(set(NomesTodasChaves_aux[i]) & set(NomesTodasChaves_aux[j]))
                    if ChavesComuns:
                        for k in range (len(ChavesComuns)):
                            Aleatorio = np.random.randint(2) 
                            if (Aleatorio == 0): #se 0 a chave fica na malha i, se 1 fica na malha j
                                NomesTodasChaves_aux[j].remove(ChavesComuns[k])
                            elif (Aleatorio == 1):
                                NomesTodasChaves_aux[i].remove(ChavesComuns[k])
                        ChavesComuns.clear()
        
        for i in range (len(NomesTodasChaves_aux)):
            N_Chaves_Malha.append(len(NomesTodasChaves_aux[i]))
        
        N_Malhas = len(NomesTodasChaves_aux)
        
        
        NomesChaves_aux = sum(NomesTodasChaves_aux, [])
        
        
        return NomesChaves_aux, N_Chaves_Malha, N_Malhas
    
    #esta função organiza as chaves levando em consideração uma solução S.
    #desta maneira as malhas/chaves são reorganizadas obrigando cada chave aberta de S a pertencer a apenas uma malha
    #desta maneira, após a reorganização cada malha terá uma chave aberta (um 0)
    def OrganizaChaves_com_S_ch (NomesTodasChaves, S_ch):

        #print ("NomesTodasChaves", NomesTodasChaves)
        
        N_Chaves_Malha = []
    
        #soma = len(NomesTodasChaves[0]) + len(NomesTodasChaves[1])+len(NomesTodasChaves[2])+len(NomesTodasChaves[3]) + len(NomesTodasChaves[4])
        
        NomesTodasChaves_aux = copy.deepcopy(NomesTodasChaves)
        #soma = len(NomesTodasChaves_aux[0])+len(NomesTodasChaves_aux[1])+len(NomesTodasChaves_aux[2])+len(NomesTodasChaves_aux[3])+len(NomesTodasChaves_aux[4])
        #print ("SOMA NOMESTODASCHAVES ANTES", soma)
        
        #print ("LEN NOMES TODAS CHAVES 1", len(NomesTodasChaves_aux))
        for i in range (len(S_ch)):
            for j in range (len(NomesTodasChaves_aux)):
                if (i!=j):
                    if S_ch[i] in NomesTodasChaves_aux[j]:
                        #print ("S_ch", S_ch)
                        #print ("S_ch[i]", S_ch[i])
                        NomesTodasChaves_aux[j].remove(S_ch[i])
        
        #soma = len(NomesTodasChaves_aux[0])+len(NomesTodasChaves_aux[1])+len(NomesTodasChaves_aux[2])+len(NomesTodasChaves_aux[3])+len(NomesTodasChaves_aux[4])
        #print ("SOMA NOMESTODASCHAVES DEPOIS", soma)
        #print ("LEN NOMES TODAS CHAVES 2", len(NomesTodasChaves_aux))
        #lembrando que ChavesMalha nessa função é o array com todas as chaves de todas as malhas (incluindo as repetidas)
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
        
        #print ("LEN NOMES TODAS CHAVES 3", len(NomesTodasChaves_aux))
        
        N_Malhas = len(NomesTodasChaves_aux)
        
        NomesChaves_aux = sum(NomesTodasChaves_aux, [])
        
        #print ("LEN NOMES CHAVES", len(NomesChaves_aux))
        
        return NomesChaves_aux, N_Chaves_Malha, N_Malhas
    
        
    #esta função organiza o Fit em ordem crescente e consequentemente os Vizinhos e a lista de movimentos 
    #em ordem crescente de acordo com o Fit já ordenado 
    def SortFitVizS_VizS_ListaMovimentos(FitVizS, N_Malhas, VizS_bin, VizS_ch):
        
        Fit_Sorted = sorted(FitVizS, reverse=False)
        
        VizS_bin_Sorted = []
        #movimentos_Sorted = []
        VizS_ch_Sorted = []
        
        for i in range (len(FitVizS)):
            index = FitVizS.index(Fit_Sorted[i])
            VizS_bin_Sorted.append(VizS_bin[index])
            #movimentos_Sorted.append(lista_movimentos_separados[index])
            VizS_ch_Sorted.append(VizS_ch[index])
            
        
        return Fit_Sorted, VizS_bin_Sorted, VizS_ch_Sorted
            
    #esta função cria um novo S e uma nova lista tabu. ela analisa malha a malha se 
    #em nenhum momento o vizinho de melhor fit é tabu com o tabu do S anterior.
    #caso uma malha seja tabu, a função pega o movimento do próximo vizinho de melhor fit (segundo, terceiro etc)
    #de forma a criar um NovoS totalmente renovado em relação ao anterior.
    def S_gerado (vizinhos_ch, vizinhos_bin, fit_vizinhos, N_Chaves_Malha, lista_tabu_completa):
        
        lista_tabu_completa_joined = [j for i in lista_tabu_completa for j in i]
        #novo_tabu = []
        novo_S_ch = []
        for i in range (len(vizinhos_ch)):
            for j in range (len(vizinhos_ch[i])):
                if (lista_tabu_completa_joined.count(vizinhos_ch[j][i])==0):
                    novo_S_ch.append(vizinhos_ch[j][i])
                    break
                    
        #novo_S = [1]*sum(N_Chaves_Malha)
        #for i in novo_tabu:
        #    novo_S[i] = 0
        
        novo_S_bin = CodificadorChaves(novo_S_ch, NomesChaves)
        fit = FitnessIndividuo (novo_S_bin, NomesChaves)
        if (Radialidade()):
            return novo_S_ch, novo_S_bin, fit
        else:
            return vizinhos_ch[0], vizinhos_bin[0], fit_vizinhos[0]
            
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
    def Vizinhos_random (S, N_Malhas, N_Chaves_Malha):
        
        Vizinhos = []
        inicio = 0
        #changes_list = [] #lista dos índices que foram alterados (para lista tabu)
        Vizinhos2 = []
        Vizinhos3_bin=[]
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
                #changes_list.append(x)
                inicio = inicio + N_Chaves_Malha[j] #atualiza o range de onde sortear a chave (por causa da malha)
                    
            fit = FitnessIndividuo(Vizinho, NomesChaves) #calcula o fit para que seja possível verificar a radialidade também
            if (Radialidade()):
                Vizinhos2.append(np.array(Vizinho))
                Vizinhos.append(Vizinho)
                FitVizinhos3.append(fit)
                Vizinhos3_bin.append(Vizinhos2[i].tolist())
                i+=1
            #else:
                #changes_list = changes_list[:len(changes_list)-N_Malhas]
        
        Vizinhos3_ch = []
        for i in range (len(Vizinhos3_bin)):
            Vizinhos3_ch.append(DecodificadorChaves(Vizinhos3_bin[i], NomesChaves))
            
        #changes_list_separated = np.array_split(changes_list, N_Malhas)
                                          
        return Vizinhos3_bin, Vizinhos3_ch, FitVizinhos3
    
    
    #vizinhos que pega o S atual e anda pra direita ou esquerda aleatoriamente o 0
    def Vizinhos_LR_v1 (S, N_Malhas, N_Chaves_Malha):
    
        random.seed(datetime.now())
        direcao = random.randint(0, 1)   #se direcao == 0, esquerda, se direcao == 1, direita
        Vizinho = S
        changes_list = []
        Vizinhos2 = []
        Vizinhos = []
        Vizinhos3_bin =[]
        FitVizinhos3 = []
        i = 0
        
        if (direcao==1):
            while i < N_Malhas:
            #for i in range(N_Malhas): #for dos vizinhos
                inicio = 0
                for j in range (N_Malhas): #for das malhas
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
                
                fit = FitnessIndividuo(Vizinho, NomesChaves) #calcula o fit para que seja possível verificar a radialidade também
                if (Radialidade()):
                    Vizinhos2.append(np.array(Vizinho))
                    Vizinhos.append(Vizinho)
                    FitVizinhos3.append(fit)
                    Vizinhos3_bin.append(Vizinhos2[i].tolist())
                    i+=1
                else:
                    changes_list = changes_list[:len(changes_list)-N_Malhas]
                    
                #Vizinhos2.append(np.array(Vizinho))
                #Vizinhos.append(Vizinho)
        elif (direcao==0):
            while i < N_Malhas:
            #for i in range(N_Malhas): #for dos vizinhos
                inicio = 0
                for j in range (N_Malhas): #for das malhas
                    index = Vizinho.index(0, inicio, inicio + N_Chaves_Malha[j])
                    if (index == inicio):         
                        Vizinho[inicio + N_Chaves_Malha[j]-1] = 0
                        Vizinho[inicio] = 1
                        changes_list.append(inicio)
                    else:
                        Vizinho[index] = 1
                        Vizinho[index-1] = 0
                        changes_list.append(index-1)            
                    inicio = inicio + N_Chaves_Malha[j]
                
                fit = FitnessIndividuo(Vizinho, NomesChaves) #calcula o fit para que seja possível verificar a radialidade também
                if (Radialidade()):
                    Vizinhos2.append(np.array(Vizinho))
                    Vizinhos.append(Vizinho)
                    FitVizinhos3.append(fit)
                    Vizinhos3_bin.append(Vizinhos2[i].tolist())
                    i+=1
                else:
                    changes_list = changes_list[:len(changes_list)-N_Malhas]
                    
                #Vizinhos2.append(np.array(Vizinho))
                #Vizinhos.append(Vizinho)
                
        #for i in range(len(Vizinhos2)):
        #    Vizinhos3.append(Vizinhos2[i].tolist())
        
        #changes_list_separated = np.array_split(changes_list, N_Malhas)
        
        Vizinhos3_ch = []
        for i in range (len(Vizinhos3_bin)):
            Vizinhos3_ch.append(DecodificadorChaves(Vizinhos3_bin[i], NomesChaves))
        
                                          
        return Vizinhos3_bin, Vizinhos3_ch, FitVizinhos3
    
    def AbreUmaChave (chave, NomesChaves):
        
        S_bin = [1] * len(NomesChaves)
        index = NomesChaves.index(chave)
        S_bin[index] = 0
        
        return S_bin
            
    
    def Vizinhos_metodoYuri(S_ch, N_Malhas, N_Chaves_Malha):
            
        NomesTodasChaves_aux = copy.deepcopy(NomesTodasChaves)                
        Vizinho = []
        Vizinhos_ch = []
        potencias_ativas_trifasicas = []
        potencias_malha_aux = []
        potencias_aux_sorted = []
        FitVizinhos3 = []
        v = 0
        for v in range (3): #for dos vizinhos (3 vizinhos)
            potencias_sorted = []
            ChavesComuns = []
            ChavesNaoComunsMalhaM =[]
            potencias_aux_sorted = []
            potencias_aux = []
            Vizinho = []
            for m in range (N_Malhas): #for de cada uma das malhas em que será realizado o fluxo
                S_umachave = AbreUmaChave(S_ch[m], NomesChaves)
                fit = FitnessIndividuo(S_umachave, NomesChaves)
                #print ("Trabalhando com a chave", S_ch[m])
                #print ("S atual: ", DecodificadorChaves(S_umachave, NomesChaves))
                #print (perdas)
                #como o fluxo foi rodado, agora é possível obter as perdas pelo código a seguir:
                potencias_ativas_trifasicas = []
            #essa parte obtém o fluxo de cada uma das linhas e armazena no vetor potencias_ativas_trifasicas
            #é válido comentar que esse vetor corresponde com o NomesChaves
                for i in range (len(NomesTodasChaves_aux)): #for de cada uma das malhas
                    potencias_malha_aux = []
                    for j in range (len(NomesTodasChaves_aux[i])): #for das chaves dentro de cada malha
                        num_linha = int(NomesTodasChaves_aux[i][j][1:]) #pega o numero da linha a que se refere
                        objeto.ativa_elemento("Line.L"+str(num_linha)) #ativa a determinada linha
                        potencias = (objeto.get_potencias_elemento()) #pega todas as potencias da linha
                        potencias_malha_aux.append(abs(potencias[0] + potencias[2] + potencias[4])) #soma somente as potências ativas entrando na linha
                    potencias_ativas_trifasicas.append(potencias_malha_aux)
                
                
                potencias_aux = []
                
                #print ("as potencias para a malha", m, "são: ", potencias_ativas_trifasicas[m])
                #print ("e as chaves são: ", NomesTodasChaves_aux[m])
        
                potencias_sorted = sorted(potencias_ativas_trifasicas[m], reverse=False) #ordena os fluxos de potência de forma crescente para uma determinada malha
                menor_fluxo_chave_malha = potencias_sorted[v+1] 
                index_menor_fluxo_chave_malha = potencias_ativas_trifasicas[m].index(menor_fluxo_chave_malha) #obtém o index do menor fluxo no vetor original potencias_ativas_trifasicas
                chave = NomesTodasChaves_aux[m][index_menor_fluxo_chave_malha] #obtém a chave no vetor NomesChaves
                
                for c in range(len(Vizinho)):
                    ChavesComuns = list(set(NomesTodasChaves_aux[m]) & set(NomesTodasChaves_aux[c])) #chaves comuns entre a malha M e a malha C (das chaves anteriores)
                    if ((chave in ChavesComuns) and (Vizinho[c] in ChavesComuns)):
                        ChavesNaoComunsMalhaM = [x for x in NomesTodasChaves_aux[m] if x not in ChavesComuns] #obtém as chaves que não são comuns entre a malha I e a malha da chave em comum
                        potencias_aux = []
                        for j in range (len(ChavesNaoComunsMalhaM)): 
                            index_aux = NomesTodasChaves_aux[m].index(ChavesNaoComunsMalhaM[j]) #encontra o index das chaves não comuns dentro do vetor NomesChaves
                            potencias_aux.append(potencias_ativas_trifasicas[m][index_aux]) #adiciona ao vetor potencias_aux os fluxos das chaves não comuns
                            potencias_aux_sorted = sorted(potencias_aux, reverse=False)
                           
                        menor_fluxo_chave_malha = potencias_aux_sorted[v+1] #pega o menor fluxo entre as chaves não comuns
                        index_menor_fluxo_chave_malha = potencias_ativas_trifasicas[m].index(menor_fluxo_chave_malha)
                        chave = NomesTodasChaves_aux[m][index_menor_fluxo_chave_malha] #obtém no vetor NomesChaves a chave de menor fluxo
                        break
                    
                Vizinho.append(chave)
                #print ("a chave escolhida para essa malha foi", chave)
    
            fit = FitnessIndividuo_ch(Vizinho)
            if (Radialidade()):
                Vizinhos_ch.append(Vizinho)
                FitVizinhos3.append(fit)
                
            Vizinho = []
    
        Vizinhos3_bin = []
        Vizinhos3_ch = Vizinhos_ch
        
        for i in range (len(Vizinhos_ch)):
            Vizinhos3_bin.append(CodificadorChaves(Vizinhos_ch[i], NomesChaves))
            
        return Vizinhos3_bin, Vizinhos3_ch, FitVizinhos3
    
    def FitnessIndividuo_ch(Ind_ch):
        for i in range (len(NomesChaves)):
            objeto.switch_elemento(NomesChaves[i], 1)
        
        for i in range(len(Ind_ch)):
            objeto.switch_elemento(Ind_ch[i], 0)
            
        objeto.solve_DSS_snapshot()       
        return objeto.get_line_losses()
    
    def Vizinhos_LR_v3 (S_ch, N_Malhas, N_Chaves_Malha):
        
        Vizinhos = []
        indexfinal = 0
        FitVizinhos = []
        
        N_Chaves_Malha_index = []
        soma = 0
        for i in range (len(N_Chaves_Malha)):
            soma = N_Chaves_Malha[i] + soma
            N_Chaves_Malha_index.append(soma - 1)
        
        
        for v in range (N_Malhas): #while dos vizinhos
            Vizinho = S_ch.copy()
            direcao = random.randint(0, 1) #1 pra direita e 0 pra esquerda
            if direcao == 0:
                direcao = -1
    
            index = NomesChaves.index(Vizinho[v])
            novoindex = index + direcao
    
            if (novoindex > N_Chaves_Malha_index[v]):
                Vizinho[v] = NomesChaves[novoindex - N_Chaves_Malha[v]]
            elif (novoindex < (N_Chaves_Malha_index[v]-N_Chaves_Malha[v])):
                Vizinho[v] = NomesChaves[N_Chaves_Malha_index[v]]
            else:
                Vizinho[v] = NomesChaves[novoindex]
                    
            fit = FitnessIndividuo_ch(Vizinho)
            #print (Vizinho)
            
            if Radialidade():
                #print ("radial")
                #print (fit)
                Vizinhos.append(Vizinho)
                FitVizinhos.append(fit)
            Vizinho = []
            
        #print (NomesChaves)    
        #print ("Vizinhos LR", Vizinhos)
        Vizinhos_bin = []
        for i in range (len(Vizinhos)):
            Vizinhos_bin.append(CodificadorChaves(Vizinhos[i], NomesChaves))
            
        return Vizinhos_bin, Vizinhos, FitVizinhos
    
    def Vizinhos_LR_v5 (S_ch, N_Malhas, N_Chaves_Malha):
        
        Vizinhos = []
        indexfinal = 0
        FitVizinhos = []
        
        N_Chaves_Malha_index = []
        soma = 0
        for i in range (len(N_Chaves_Malha)):
            soma = N_Chaves_Malha[i] + soma
            N_Chaves_Malha_index.append(soma - 1)
        
        
        for v in range (N_Malhas): #while dos vizinhos
            Vizinho = S_ch.copy()
            direcao = random.randint(0, 1) #1 pra direita e 0 pra esquerda
            if direcao == 0:
                direcao = -1
    
            index = NomesChaves.index(Vizinho[v])
            novoindex = index + direcao
    
            if (novoindex > N_Chaves_Malha_index[v]):
                Vizinho[v] = NomesChaves[novoindex - N_Chaves_Malha[v]]
            elif (novoindex < (N_Chaves_Malha_index[v]-N_Chaves_Malha[v])):
                Vizinho[v] = NomesChaves[N_Chaves_Malha_index[v]]
            else:
                Vizinho[v] = NomesChaves[novoindex]
                    
            fit = FitnessIndividuo_ch(Vizinho)
            #print (Vizinho)
            
            if Radialidade():
                #print ("radial")
                #print (fit)
                Vizinhos.append(Vizinho)
                FitVizinhos.append(fit)
            Vizinho = []
            
        for v in range (N_Malhas): #while dos vizinhos
            Vizinho = S_ch.copy()
            direcao = random.randint(0, 1) #1 pra direita e 0 pra esquerda
            if direcao == 0:
                direcao = -2
            else:
                direcao = 2
    
            index = NomesChaves.index(Vizinho[v])
            novoindex = index + direcao
    
            if (novoindex > N_Chaves_Malha_index[v]):
                Vizinho[v] = NomesChaves[novoindex - N_Chaves_Malha[v]]
            elif (novoindex < (N_Chaves_Malha_index[v]-N_Chaves_Malha[v])):
                Vizinho[v] = NomesChaves[N_Chaves_Malha_index[v]]
            else:
                Vizinho[v] = NomesChaves[novoindex]
                    
            fit = FitnessIndividuo_ch(Vizinho)
            #print (Vizinho)
            
            if Radialidade():
                #print ("radial")
                #print (fit)
                Vizinhos.append(Vizinho)
                FitVizinhos.append(fit)
            Vizinho = []
            
        #print (NomesChaves)    
        #print ("Vizinhos LR", Vizinhos)
        Vizinhos_bin = []
        for i in range (len(Vizinhos)):
            Vizinhos_bin.append(CodificadorChaves(Vizinhos[i], NomesChaves))
            
        return Vizinhos_bin, Vizinhos, FitVizinhos
    
    def Vizinhos_LR_v6 (S_ch, N_Malhas, N_Chaves_Malha, N_vizinhos):
        
        Vizinhos = []
        indexfinal = 0
        FitVizinhos = []
        
        N_Chaves_Malha_index = []
        soma = 0
        for i in range (len(N_Chaves_Malha)):
            soma = N_Chaves_Malha[i] + soma
            N_Chaves_Malha_index.append(soma - 1)
        
        
        for v in range (N_Malhas): #while dos vizinhos
            Vizinho = S_ch.copy()
            direcao = random.randint(0, 1) #1 pra direita e 0 pra esquerda
            if direcao == 0:
                direcao = -1
    
            index = NomesChaves.index(Vizinho[v])
            novoindex = index + direcao
    
            if (novoindex > N_Chaves_Malha_index[v]):
                Vizinho[v] = NomesChaves[novoindex - N_Chaves_Malha[v]]
            elif (novoindex < (N_Chaves_Malha_index[v]-N_Chaves_Malha[v])):
                Vizinho[v] = NomesChaves[N_Chaves_Malha_index[v]]
            else:
                Vizinho[v] = NomesChaves[novoindex]
                    
            fit = FitnessIndividuo_ch(Vizinho)
            #print (Vizinho)
            
            if Radialidade():
                #print ("radial")
                #print (fit)
                Vizinhos.append(Vizinho)
                FitVizinhos.append(fit)
            Vizinho = []
            
        for v in range (N_Malhas): #while dos vizinhos
            Vizinho = S_ch.copy()
            direcao = random.randint(0, 1) #1 pra direita e 0 pra esquerda
            if direcao == 0:
                direcao = -2
            else:
                direcao = 2
    
            index = NomesChaves.index(Vizinho[v])
            novoindex = index + direcao
    
            if (novoindex > N_Chaves_Malha_index[v]):
                Vizinho[v] = NomesChaves[novoindex - N_Chaves_Malha[v]]
            elif (novoindex < (N_Chaves_Malha_index[v]-N_Chaves_Malha[v])):
                Vizinho[v] = NomesChaves[N_Chaves_Malha_index[v]]
            else:
                Vizinho[v] = NomesChaves[novoindex]
                    
            fit = FitnessIndividuo_ch(Vizinho)
            #print (Vizinho)
            
            if Radialidade():
                #print ("radial")
                #print (fit)
                Vizinhos.append(Vizinho)
                FitVizinhos.append(fit)
            Vizinho = []
            
        for v in range (N_Malhas): #while dos vizinhos
            Vizinho = S_ch.copy()
            direcao = random.randint(0, 1) #1 pra direita e 0 pra esquerda
            if direcao == 0:
                direcao = -3
            else:
                direcao = 3
    
            index = NomesChaves.index(Vizinho[v])
            novoindex = index + direcao
    
            if (novoindex > N_Chaves_Malha_index[v]):
                Vizinho[v] = NomesChaves[novoindex - N_Chaves_Malha[v]]
            elif (novoindex < (N_Chaves_Malha_index[v]-N_Chaves_Malha[v])):
                Vizinho[v] = NomesChaves[N_Chaves_Malha_index[v]]
            else:
                Vizinho[v] = NomesChaves[novoindex]
                    
            fit = FitnessIndividuo_ch(Vizinho)
            #print (Vizinho)
            
            if Radialidade():
                #print ("radial")
                #print (fit)
                Vizinhos.append(Vizinho)
                FitVizinhos.append(fit)
            Vizinho = []
            
        #print (NomesChaves)    
        #print ("Vizinhos LR", Vizinhos)
        Vizinhos_bin = []
        for i in range (len(Vizinhos)):
            Vizinhos_bin.append(CodificadorChaves(Vizinhos[i], NomesChaves))
            
        return Vizinhos_bin, Vizinhos, FitVizinhos
    
    def Vizinhos_LR_v4 (S_ch, N_Malhas, N_Chaves_Malha):
        
        Vizinhos = []
        indexfinal = 0
        FitVizinhos = []
        
        N_Chaves_Malha_index = []
        soma = 0
        for i in range (len(N_Chaves_Malha)):
            soma = N_Chaves_Malha[i] + soma
            N_Chaves_Malha_index.append(soma - 1)
        
        
        for v in range (N_Malhas): #while dos vizinhos
            Vizinho = S_ch.copy()
            direcao = random.randint(0, 1) #1 pra direita e 0 pra esquerda
            direcao == -1
                
            index = NomesChaves.index(Vizinho[v])
            novoindex = index + direcao
    
            if (novoindex > N_Chaves_Malha_index[v]):
                Vizinho[v] = NomesChaves[novoindex - N_Chaves_Malha[v]]
            elif (novoindex < (N_Chaves_Malha_index[v]-N_Chaves_Malha[v])):
                Vizinho[v] = NomesChaves[N_Chaves_Malha_index[v]]
            else:
                Vizinho[v] = NomesChaves[novoindex]
                    
            fit = FitnessIndividuo_ch(Vizinho)
            #print (Vizinho)
            
            if Radialidade():
                #print ("radial")
                #print (fit)
                Vizinhos.append(Vizinho)
                FitVizinhos.append(fit)
            Vizinho = []
            
        for v in range (N_Malhas): #while dos vizinhos
            Vizinho = S_ch.copy()
            direcao = random.randint(0, 1) #1 pra direita e 0 pra esquerda
            direcao = 1
    
            index = NomesChaves.index(Vizinho[v])
            novoindex = index + direcao
    
            if (novoindex > N_Chaves_Malha_index[v]):
                Vizinho[v] = NomesChaves[novoindex - N_Chaves_Malha[v]]
            elif (novoindex < (N_Chaves_Malha_index[v]-N_Chaves_Malha[v])):
                Vizinho[v] = NomesChaves[N_Chaves_Malha_index[v]]
            else:
                Vizinho[v] = NomesChaves[novoindex]
                    
            fit = FitnessIndividuo_ch(Vizinho)
            #print (Vizinho)
            
            if Radialidade():
                #print ("radial")
                #print (fit)
                Vizinhos.append(Vizinho)
                FitVizinhos.append(fit)
            Vizinho = []
            
        #print (NomesChaves)    
        #print ("Vizinhos LR", Vizinhos)
        Vizinhos_bin = []
        for i in range (len(Vizinhos)):
            Vizinhos_bin.append(CodificadorChaves(Vizinhos[i], NomesChaves))
            
        return Vizinhos_bin, Vizinhos, FitVizinhos
    

        

            
    
    def Vizinhos_LR_v2 (S, N_Malhas, N_Chaves_Malha):
    
        random.seed(datetime.now())
        
        Vizinho = S
        changes_list = []
        Vizinhos2 = []
        Vizinhos = []
        Vizinhos3_bin =[]
        FitVizinhos3 = []
        i = 0
        
        while i < N_Malhas: #while dos vizinhos
            
            #desta maneira vai variar esquerda e direita em cada vizinho
        #for i in range(N_Malhas): #for dos vizinhos
            inicio = 0
            for j in range (N_Malhas): #for das malhas
                direcao = random.randint(0, 1)   #se direcao == 0, esquerda, se direcao == 1, direita
                if direcao == 0:
                    andar = 1
                else:
                    andar = -1
                index = Vizinho.index(0, inicio, inicio + N_Chaves_Malha[j])
                if (index == (inicio + N_Chaves_Malha[j]-1)):         
                    Vizinho[index] = 1
                    Vizinho[inicio] = 0
                    changes_list.append(inicio)
                else:
                    Vizinho[index] = 1
                    Vizinho[index+1] = 0
                    changes_list.append(index+andar)            
                inicio = inicio + N_Chaves_Malha[j]
            
            fit = FitnessIndividuo(Vizinho, NomesChaves) #calcula o fit para que seja possível verificar a radialidade também
            if (Radialidade()):
                Vizinhos2.append(np.array(Vizinho))
                Vizinhos.append(Vizinho)
                FitVizinhos3.append(fit)
                Vizinhos3_bin.append(Vizinhos2[i].tolist())
                i+=1
            else:
                changes_list = changes_list[:len(changes_list)-N_Malhas]
        
        
        Vizinhos3_ch = []
        for i in range (len(Vizinhos3_bin)):
            Vizinhos3_ch.append(DecodificadorChaves(Vizinhos3_bin[i], NomesChaves))
        
                                          
        return Vizinhos3_bin, Vizinhos3_ch, FitVizinhos3
    
    def Vizinhos_PFO (S, N_Malhas, N_Chaves_Malha):
        
        S_fechado = sum(N_Chaves_Malha)*[1]
        perdas = FitnessIndividuo(S_fechado, NomesChaves)
        print (perdas)
            
        potencias_ativas_trifasicas = []
        potencias_malha_aux = []
        
        #essa parte obtém o fluxo de cada uma das linhas e armazena no vetor potencias_ativas_trifasicas
        #é válido comentar que esse vetor corresponde com o NomesChaves
        for i in range (len(NomesTodasChaves)): #for de cada uma das malhas
            potencias_malha_aux = []
            for j in range (len(NomesTodasChaves[i])): #for das chaves dentro de cada malha
                num_linha = int(NomesTodasChaves[i][j][1:]) #pega o numero da linha a que se refere
                objeto.ativa_elemento("Line.L"+str(num_linha)) #ativa a determinada linha
                potencias = (objeto.get_potencias_elemento()) #pega todas as potencias da linha
                potencias_malha_aux.append(abs(potencias[0] + potencias[2] + potencias[4])) #soma somente as potências ativas entrando na linha
            potencias_ativas_trifasicas.append(potencias_malha_aux)
        
        S = []
        Vizinhos_ch = []
        potencias_aux = []
        
        k = 0
        #for k in range (N_Malhas):
        while k < N_Malhas:
            for i in range (len(potencias_ativas_trifasicas)): #for de cada uma das malhas
                potencias_sorted = sorted(potencias_ativas_trifasicas[i], reverse=False) #ordena os fluxos de potência de forma crescente para uma determinada malha
                menor_fluxo_chave_malha = potencias_sorted[k]
                index_menor_fluxo_chave_malha = potencias_ativas_trifasicas[i].index(menor_fluxo_chave_malha) #obtém o index do menor fluxo no vetor original potencias_ativas_trifasicas
                chave = NomesTodasChaves[i][index_menor_fluxo_chave_malha] #obtém a chave no vetor NomesChaves
                
                if (chave in S): #se a chave já tiver sido escolhida 
                    index_malhacomum = S.index(chave) #descobre a malha que a chave selecionada está em comum
                    ChavesComuns = [value for value in NomesTodasChaves[i] if value in NomesTodasChaves[index_malhacomum]] #obtém as chaves em comum entre a malha I e a malha da chave em comum
                    ChavesNaoComunsMalhai = [x for x in NomesTodasChaves[i] if x not in ChavesComuns] #obtém as chaves que não são comuns entre a malha I e a malha da chave em comum
                    potencias_aux = []
                    for j in range (len(ChavesNaoComunsMalhai)): 
                        index_aux = NomesTodasChaves[i].index(ChavesNaoComunsMalhai[j]) #encontra o index das chaves não comuns dentro do vetor NomesChaves
                        potencias_aux.append(potencias_ativas_trifasicas[i][index_aux]) #adiciona ao vetor potencias_aux os fluxos das chaves não comuns
                        potencias_aux_sorted = sorted(potencias_aux, reverse=False) #ordena as potencias em ordem crescente
                
                    menor_fluxo_chave_malha = potencias_aux_sorted[k] #pega o menor fluxo entre as chaves não comuns
                    index_menor_fluxo_chave_malha = potencias_ativas_trifasicas[i].index(menor_fluxo_chave_malha)
                    chave = NomesTodasChaves[i][index_menor_fluxo_chave_malha] #obtém no vetor NomesChaves a chave de menor fluxo
                S.append(chave)
            
            fit = FitnessIndividuo (S, NomesChaves)
            if Radialidade():
                Vizinhos_ch.append(S)
                k+=1
            S = []
        
        Vizinhos3_bin = []
        Vizinhos3_ch = Vizinhos_ch
        FitVizinhos3 = []
        for i in range (len(Vizinhos_ch)):
            Vizinhos3_bin.append(CodificadorChaves(Vizinhos_ch[i], NomesChaves))
            FitVizinhos3.append(FitnessIndividuo(Vizinhos3_bin[i], NomesChaves))
        
        print (FitVizinhos3)
        
        return Vizinhos3_bin, Vizinhos3_ch, FitVizinhos3
        
     
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

        #NomesChaves, N_Chaves_Malha, N_Malhas = OrganizaChaves(NomesTodasChaves)
        
        #S_ch = S_inicial(NomesTodasChaves) #obtém o S inicial pelo método do menor fluxo nas malhas
        #S_ch = ['S33','S37','S35','S36','S34']
        #S_ch = ['S72','S69','S71','S70','S73']
        S_ch = ["S96","S88","S91","S87","S89","S94","S95","S93","S92","S90","S84","S85","S86"]
        NomesChaves, N_Chaves_Malha, N_Malhas = OrganizaChaves_com_S_ch (NomesTodasChaves, S_ch)
        #S_bin = GeraS(N_Chaves_Malha, N_Malhas)
        #S_ch = DecodificadorChaves(S_bin, NomesChaves)
        
        S_bin = CodificadorChaves(S_ch, NomesChaves)
               
        FitS = FitnessIndividuo(S_bin, NomesChaves)
        print ("Configuração inicial: ", S_ch)
        
        BestFit = FitS
        BestS_bin = S_bin
        BestS_ch = S_ch
        Iter = 0
        
        status_tensao, min_tensao, max_tensao = Limites_tensao(6581)
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
         #tamanho da lista tabu
        
        
        lista_tabu = [-1]*sum(N_Chaves_Malha)
        novo_tabu = [-1]*sum(N_Chaves_Malha)
        ChavesAbertasVizS = []
        lista_tabu_completa = []
        
        melhorvizinho = 0
        vizinhogerado = 0
        TodosS = []
        
        aux1 = [0]
        aux2 = [0]
        aux3 = [0]
                                
        while (i-BestIter) <= BTMax:
            
            i+=1
            #print ("================NOVA ITERAÇÃO=================", i)
            #print ("Partindo com Solução Corrente:", S_ch)
            
            
            #if (i%3 == 0):
            #aux1[0], aux2[0], aux3[0] = Vizinhos_metodoYuri(S_ch, N_Malhas, N_Chaves_Malha) 
            #aux1[0], aux2[0], aux3[0] = Vizinhos_random(S_bin, N_Malhas, N_Chaves_Malha)
            """aux1[0], aux2[0], aux3[0] = Vizinhos_LR_v3(S_ch, N_Malhas, N_Chaves_Malha)
            #aux1[0], aux2[0], aux3[0] = Vizinhos_LR_v4(S_ch, N_Malhas, N_Chaves_Malha)
            
            VizS_bin = [item for sublist in aux1 for item in sublist]
            VizS_ch = [item for sublist in aux2 for item in sublist]
            FitVizS = [item for sublist in aux3 for item in sublist]"""
            
            VizS_bin, VizS_ch, FitVizS = Vizinhos_LR_v3 (S_ch, N_Malhas, N_Chaves_Malha)
            
            lista_tabu = novo_tabu    
            
            #vai adicionando os movimentos de cada S selecionado à lista_tabu_completa;
            #conforme ela vai preenchendo os movimentos mais antigos vão sendo eliminados
            if (len(lista_tabu_completa)==T): 
                lista_tabu_completa.pop(0)
                lista_tabu_completa.append(lista_tabu)
            else:
                lista_tabu_completa.append(lista_tabu)
                
            #organiza o fit, os movimentos e os vizinhos de acordo com o fit (do melhor pro pior)    
            fitorg, vizorg_bin, vizorg_ch = SortFitVizS_VizS_ListaMovimentos(FitVizS, N_Malhas, VizS_bin, VizS_ch)
            
            #caso já tenha passado a primeira iteração já é possível calcular o S gerado (levando em consideração a LT)
            """if (i>1):
                S_gerado_ch, S_gerado_bin, fit_S_gerado = S_gerado(vizorg_ch, vizorg_bin, fitorg, N_Chaves_Malha, lista_tabu_completa)
                VizS_bin.append(S_gerado_bin)
                VizS_ch.append(S_gerado_ch)
                FitVizS.append(fit_S_gerado)
                print ("S gerado: ", S_gerado_ch)"""
            
            #reorganiza novamente agora com o novo vizinho (S gerado)
            fitorg, vizorg_bin, vizorg_ch = SortFitVizS_VizS_ListaMovimentos(FitVizS, N_Malhas, VizS_bin, VizS_ch)
            
            #for j in range(len(vizorg_ch)):
            #   print ("Vizinho", j, vizorg_ch[j], "Fit:", fitorg[j])
        
            #PARTE DA ESCOLHA DO NOVO S:
            if (fitorg[0] < BestFit): #função de aspiração: caso o melhor vizinho seja a melhor solução já encontrada mas estiver na LT, aceita mesmo assim 
                S_bin = vizorg_bin[0]
                S_ch = vizorg_ch[0]
                fit_novo_S = fitorg[0]
                novo_tabu = S_ch
            else: #caso esteja na lista tabu, entra pro processo de escolha
                lista_tabu_completa_flat = [item for sublist in lista_tabu_completa for item in sublist]
                for m in range (len(vizorg_ch)):
                    if (any(n in lista_tabu_completa_flat for n in vizorg_ch[m]) == False):
                        #print ("O vizinho escolhido foi o", m)
                        S_bin = vizorg_bin[m]
                        S_ch = vizorg_ch[m]
                        fit_novo_S = fitorg[m]
                        novo_tabu = S_ch
                        break
            
            #se o fit do novo S for o melhor já encontrado:
            if (fit_novo_S < BestFit):
                
                BestS_bin = S_bin
                BestS_ch = S_ch
                BestIter = i
                #print ("Novos BestFit e BestS encontrados")
                BestFit =  FitnessIndividuo_ch(BestS_ch)
                status_tensao, min_tensao, max_tensao = Limites_tensao(6581)
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
            
            S_bin = CodificadorChaves(S_ch, NomesChaves)
            #NomesChaves, N_Chaves_Malha, N_Malhas = OrganizaChaves(NomesTodasChaves)
            
            VizS_bin.clear()
            VizS_ch.clear()
            FitVizS.clear()


            fitorg.clear()
            FitVizS.clear()
            vizorg_bin.clear()
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
            end = time.time()
            TempoGasto.append(end-start)
            MelhoresFitsGerais.append(BestFit)
            MelhoresSGerais.append(BestS_ch)
            MinimasTensoesGerais.append(min_tensao)
            MaximasTensoesGerais.append(max_tensao)
            solucao_otima = TestaSolucaoOtima(BestS_ch, N_Barras_Sistema)
            if solucao_otima == 1:
                solucoesotimas += 1
                print ("Solução ótima encontrada!")
            
        NomesChaves = []
        NomesTodasChaves = []

    
    
print ("================RESULTADO FINAL================")
for i in range (len(MelhoresFitsGerais)):
    print ("Config: ", MelhoresSGerais[i])
    print ("Fit: ", round (MelhoresFitsGerais[i],3), "kW")
    print ("Tensao minima: ", round (MinimasTensoesGerais[i],4), "pu")           
    print ("Tensão máxima: ", round (MaximasTensoesGerais[i],4), "pu")
    print ("Tempo gasto: ", TempoGasto[i])
    
    
print ("Média de perdas para",N_Vezes,"configurações obtidas para sistema de", 
       N_Barras_Sistema, "barras: ", round(sum(MelhoresFitsGerais)/len(MelhoresFitsGerais),3))

print ("O sistema obteve", solucoesotimas,"soluções ótimas para", len(MelhoresFitsGerais), "soluções entregues")   

print ("Tempo gasto total: ", sum(TempoGasto))

    
            
    
    
    


