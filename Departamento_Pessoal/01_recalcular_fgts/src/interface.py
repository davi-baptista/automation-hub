from PySimpleGUI import PySimpleGUI as sg

def input_filter(key, value):
    if value.isdigit() or value == "":
        if (key == 'cnpj' or key == 'cnpj_editar') and len(value) > 14:
            return value[:-1]
        if (key == 'pis' or key == 'pis_editar') and len(value) > 11:
            return value[:-1]
        if (key == 'competencia' or key == 'competencia_editar') and len(value) > 6:
            return value[:-1]
        if (key == 'datapag' or key == 'datapag_editar') and len(value) > 8:
            return value[:-1]
        return value
    else:
        return value[:-1]

class Competencias():
    def __init__(self):
        self._cnpj = []
        self._pis = []
        self._cadapis = []
        self._competencia = []
        self._datapag = []
        self._certificado = []
        self._arquivo = []
        self._tipopis = []
    
    def addcompetencia(self, cnpj, competencia, datapag, certificado, arquivo, tipopis):
        self._cnpj.append(cnpj)
        self._pis.append(self._cadapis.copy())
        self._competencia.append(competencia)
        self._datapag.append(datapag)
        self._certificado.append(certificado)
        self._arquivo.append(arquivo)
        self._tipopis.append(tipopis)
        self._cadapis.clear()
        
    @property
    def cadapis(self):
        return self._cadapis
    
    @cadapis.setter
    def cadapis(self, novo_valor):
        self._cadapis = novo_valor

    def addcadapis(self, cadapis):
        self._cadapis.append(cadapis)
        
    def rmcadapis(self, indice):
        self._cadapis.pop(indice)
    
    @property
    def cnpj(self):
        return self._cnpj
    
    @cnpj.setter
    def cnpj(self, novo_valor):
        self._cnpj.append(novo_valor)
    
    def get_cnpj(self, indice):
        return self._cnpj[indice]
    
    def set_cnpj(self, indice, novo_valor):
        self._cnpj[indice] = novo_valor
    
    @property
    def pis(self):
        return self._pis
    
    @pis.setter
    def pis(self, novo_valor):
        self._pis.append(novo_valor)
    
    def get_pis(self, indice):
        return self._pis[indice]
    
    def set_pis(self, indice, novo_valor):
        self._pis[indice] = novo_valor

    @property
    def competencia(self):
        return self._competencia
    
    @competencia.setter
    def competencia(self, novo_valor):
        self._competencia.append(novo_valor)
    
    def get_competencia(self, indice):
        return self._competencia[indice]
    
    def set_competencia(self, indice, novo_valor):
        self._competencia[indice] = novo_valor

    @property
    def datapag(self):
        return self._datapag
    
    @datapag.setter
    def datapag(self, novo_valor):
        self._datapag.append(novo_valor)
    
    def get_datapag(self, indice):
        return self._datapag[indice]
    
    def set_datapag(self, indice, novo_valor):
        self._datapag[indice] = novo_valor

    @property
    def certificado(self):
        return self._certificado
    
    @certificado.setter
    def certificado(self, novo_valor):
        self._certificado.append(novo_valor)
    
    def get_certificado(self, indice):
        return self._certificado[indice]
    
    def set_certificado(self, indice, novo_valor):
        self._certificado[indice] = novo_valor

    @property
    def arquivo(self):
        return self._arquivo
    
    @arquivo.setter
    def arquivo(self, novo_valor):
        self._arquivo.append(novo_valor)
    
    def get_arquivo(self, indice):
        return self._arquivo[indice]
    
    def set_arquivo(self, indice, novo_valor):
        self._arquivo[indice] = novo_valor
        
    @property
    def tipopis(self):
        return self._tipopis
    
    @tipopis.setter
    def tipopis(self, novo_valor):
        self._tipopis.append(novo_valor)
    
    def get_tipopis(self, indice):
        return self._tipopis[indice]
    
    def set_tipopis(self, indice, novo_valor):
        self._tipopis[indice] = novo_valor
    
    def excluiratributo(self, indice):
        self._cnpj.pop(indice)
        self._pis.pop(indice)
        self._competencia.pop(indice)
        self._datapag.pop(indice)
        self._arquivo.pop(indice)
        self._certificado.pop(indice)
        self._tipopis.pop(indice)
        
def main():
    lista_pis = []
    competencias = Competencias()
    cont = 0
    layout = [
        [sg.Text('CNPJ:', size=(10, 1), justification='right'), sg.InputText(key='cnpj', enable_events=True)],
        [sg.Text('PIS', size=(10, 1), justification='right'), sg.InputText(key='pis', enable_events=True, disabled=False), sg.Button('+', key='btn', size=(2,1), visible=True)],
        [sg.Column([[sg.Radio('Adicionar PIS', 'pisopcoes', key='umpis', default=True, enable_events=True), sg.Radio('Adicionar Todos os PIS', 'pisopcoes', key='todospis', enable_events=True)]], justification='center')],
        [sg.Column([[sg.Listbox(values=lista_pis, size=(20, 3), key='listbox', visible=True)]], justification='center'), sg.Button('Excluir', visible=True)],
        [sg.Text('Compentencia', size=(10, 1), justification='right'), sg.InputText(key='competencia', enable_events=True)],
        [sg.Text('Data pag.:', size=(10, 1), justification='right'), sg.InputText(key='datapag', enable_events=True)],
        [sg.Text('Cerficado:', size=(10, 1), justification='right'), sg.InputText('653c23111645c132', key='certificado', enable_events=True)],
        [sg.Text('Arquivo:', size=(10, 1), justification='right'), sg.InputText(key='arquivo', enable_events=True), sg.FolderBrowse()],
        [sg.Button('Enviar competencia'), sg.Button('Adicionar competencia'), sg.Button('Historico'), sg.Button('Sair')]
    ]
    janela = sg.Window('Robo FGTS', layout)
    enviarpis = True
    while True:
        eventos, valores = janela.read()
        if eventos == sg.WINDOW_CLOSED or eventos == 'Sair':
            break
        
        elif eventos == 'Enviar competencia':
            competencia = valores['competencia']
            if len(competencia) != 0:
                sg.popup_error('Campo *Competencia contem um valor a ser enviado.')
                continue
            if len(competencias.competencia) == 0:
                sg.popup_error('Nenhuma competencia adicionada.')
                continue
            sg.popup("Competência enviada para automação!")
            for i in range(len(competencias.cnpj)):
                Bot.adicionarlista(Bot, competencias.get_cnpj(i), competencias.get_pis(i), competencias.get_competencia(i), competencias.get_datapag(i), competencias.get_certificado(i), competencias.get_arquivo(i), competencias.get_tipopis(i))
            pyautogui.hotkey('win', 'm')
            Bot.main()
            
        elif eventos == 'btn':
            pis = valores['pis']
            if len(pis) < 11:
                sg.popup('Campo *PIS não tem dígitos suficientes.')
                continue
            if pis in competencias.cadapis:
                sg.popup('Este *PIS já foi adicionado anteriormente.')
                continue
            cont += 1
            #-Adiciono pis no meu objeto competencias no atributo que pega cada pis
            competencias.addcadapis(pis)
            novo_item = f'PIS {str(cont)} -> {pis}'
            lista_pis.append(novo_item)
            janela['pis'].update("")
            janela['btn'].update(cont)
            janela['listbox'].update(values=lista_pis)
        
        elif eventos == 'todospis' and valores['todospis']:
            janela['pis'].update(disabled=True)
            janela['btn'].update(visible=False)
            janela['listbox'].update(visible=False)
            janela['Excluir'].update(visible=False)
            janela['pis'].update('')
            
            enviarpis = False
            
        elif eventos == 'umpis' and valores['umpis']:
            janela['pis'].update(disabled=False)
            janela['btn'].update(visible=True)
            janela['listbox'].update(visible=True)
            janela['Excluir'].update(visible=True)
            enviarpis = True
                
        elif eventos == 'Excluir':
            if valores['listbox']:
                indice_pis = lista_pis.index(valores['listbox'][0])
                lista_pis.pop(indice_pis)
                competencias.rmcadapis(indice_pis)
                nova_lista_pis = []
                aux = 0
                for i in competencias.cadapis:
                    aux += 1
                    nova_lista_pis.append(f'PIS {aux} -> {i}')
                lista_pis = nova_lista_pis
                janela['listbox'].update(values=lista_pis)
                cont -= 1
                if cont == 0:
                    janela['btn'].update('+')
                else:
                    janela['btn'].update(cont)
                    
        elif eventos == 'Adicionar competencia':
            todospis = False
            tipopis = "UM_PIS"
            cnpj = valores['cnpj']
            cadapis = valores['pis']
            competencia = valores['competencia']
            datapag = valores['datapag']
            certificado = valores['certificado']
            arquivo = valores['arquivo']
            
            #-Caso a pessoa marque a opção para adicionar todos os pis
            if valores['todospis']:
                todospis = True
                tipopis = "TODOS_PIS"
            
            if competencia in competencias.competencia:
                sg.popup('Competência já adicionada anteriormente.')
                continue
            if len(cadapis) != 0:
                sg.popup_error('Campo *PIS tem um valor a ser enviado')
                continue
            
            erros = []
            if not cnpj:
                erros.append('Campo *CNPJ está vazio.')
            if not competencia:
                erros.append('Campo *Competência está vazio.')
            if not datapag:
                erros.append('Campo *Data pag está vazio.')
            if not certificado:
                erros.append('Campo *Certificado está vazio.')
            if not arquivo:
                erros.append('Campo *Arquivo está vazio.')
            if len(competencias.cadapis) == 0 and not todospis:
                erros.append('Nenhum *PIS adicionado ainda.')
            if erros:
                sg.popup_error('\n'.join(erros))
                continue
            erros = []
            if len(cnpj) < 14:
                erros.append(f'Campo *CNPJ não tem dígitos suficientes.')
            if len(competencia) < 6:
                erros.append(f'Campo *Competência não tem dígitos suficientes.')
            if len(datapag) < 8:
                erros.append(f'Campo *Data pag não tem dígitos suficientes.')
            if erros:
                sg.popup_error('\n'.join(erros))
                continue
            
            if enviarpis == False:
                #-Excluir todos os valores encontrados na listbox
                lista_pis = []
                competencias.cadapis.clear()
                janela['pis'].update('')
                janela['listbox'].update(values=lista_pis)
                janela['btn'].update('+')
                cont = 0
            
            competencias.addcompetencia(cnpj, competencia, datapag, certificado, arquivo, tipopis)
            sg.popup("Competência adicionada na lista.")
            
            #-Atualizando tudo para voltar ao mesmo padrão de quando aberto
            cont = 0
            lista_pis = []
            janela['listbox'].update("")
            janela['btn'].update('+')
            janela['competencia'].update("")
            janela['arquivo'].update("")
            janela['umpis'].update(True)
            janela['pis'].update(disabled=False)
            janela['btn'].update(visible=True)
            janela['listbox'].update(visible=True)
            janela['Excluir'].update(visible=True)
            enviarpis = True
            
        elif eventos == 'Historico':
            if len(competencias.competencia) == 0:
                sg.popup_error('Nenhuma competência adicionada')
                continue
            historico_layout = [
                [sg.Text('Escolha uma opção:')],
                [sg.Listbox(competencias.competencia, size=(20, 3), key='opcao')],
                [sg.Button('OK'), sg.Button('Editar'), sg.Button('Excluir')]
            ]
            historico_janela = sg.Window('Escolher Opção', historico_layout, modal=True)
            while True:
                eventos_popup, valores_popup = historico_janela.read()
                if eventos_popup == sg.WINDOW_CLOSED or eventos_popup == 'OK':
                    break
                
                elif eventos_popup == 'Editar':
                    if valores_popup['opcao']:
                        indice = competencias.competencia.index(valores_popup['opcao'][0])
                        valor_condicional1 = competencias.get_tipopis(indice)
                        #-Criando novo layout para janela de edição
                        editar_layout = [
                            [sg.Text('CNPJ:', size=(10, 1), justification='right'), sg.InputText(competencias.get_cnpj(indice), key='cnpj_editar', enable_events=True)],
                            [sg.Text('PIS', size=(10, 1), justification='right'), sg.InputText(key='pis_editar', enable_events=True, disabled=valor_condicional1 != 'UMPIS'), sg.Button('+', key='btn_editar', size=(2,1), visible=valor_condicional1 == 'UM_PIS')],
                            [sg.Column([[sg.Radio('Adicionar PIS', 'pisopcoes_editar', key='umpis_editar', default=valor_condicional1 == 'UM_PIS', enable_events=True), sg.Radio('Adicionar Todos os PIS', 'pisopcoes_editar', key='todospis_editar', default=valor_condicional1 == 'TODOS_PIS', enable_events=True)]], justification='center')],
                            [sg.Column([[sg.Listbox(values=lista_pis, size=(20, 3), key='listbox_editar', visible=valor_condicional1 == 'UM_PIS')]], justification='center'), sg.Button('Excluir', visible=valor_condicional1 == 'UM_PIS')],
                            [sg.Text('Compentencia', size=(10, 1), justification='right'), sg.InputText(competencias.get_competencia(indice), key='competencia_editar', enable_events=True)],
                            [sg.Text('Data pag.:', size=(10, 1), justification='right'), sg.InputText(competencias.get_datapag(indice), key='datapag_editar', enable_events=True)],
                            [sg.Text('Cerficado:', size=(10, 1), justification='right'), sg.InputText(competencias.get_certificado(indice), key='certificado_editar', enable_events=True)],
                            [sg.Text('Arquivo:', size=(10, 1), justification='right'), sg.InputText(competencias.get_arquivo(indice), key='arquivo_editar', enable_events=True), sg.FolderBrowse()],
                            [sg.Button('Salvar'), sg.Button('Sair')]
                        ]
                        editar_janela = sg.Window('Modo de Edição', editar_layout, modal=True)
                        
                        #-Escrever lista de pis com enumeração e atualizando botão
                        editar_janela.finalize()
                        editar_janela['listbox_editar'].update(values=[f'PIS {i+1} -> {pis}' for i, pis in enumerate(competencias.get_pis(indice))])
                        cont2 = len(competencias.get_pis(indice))
                        editar_janela['btn_editar'].update(cont2)
                        
                        #-Abrindo nova janela de edição
                        apagarpis_editar = False
                        while True:
                            eventos_edicao, valores_edicao = editar_janela.read()
                            if eventos_edicao == sg.WINDOW_CLOSED or eventos_edicao == 'Sair':
                                break
                            
                            elif eventos_edicao == 'btn_editar':
                                #-Pega o valor do pis no campo de texto correspondente
                                pis = valores_edicao['pis_editar']
                                #-Confere se o tamanho do pis esta certo
                                if len(pis) < 11:
                                    sg.popup('Campo *PIS não tem dígitos suficientes.')
                                    continue
                                #-Pego os pis já adicionados anteriormente e que estão na lista atual
                                lista_pis_atual = competencias.get_pis(indice)
                                #-Se o pis ja foi adicionado na lista atual
                                if pis in lista_pis_atual and pis != valores_popup['opcao_editar'][0]:
                                    sg.popup('Este *PIS já foi adicionado anteriormente.')
                                    continue
                                #-Adiciono pis na lista atual
                                lista_pis_atual.append(pis)
                                #-Atualizo a lista de pis no meu objeto competencias
                                competencias.set_pis(indice, lista_pis_atual)
                                #-Agora atualizo a janela para mostrar atualizado
                                editar_janela['pis_editar'].update("")
                                cont2 += 1
                                editar_janela['btn_editar'].update(cont2)
                                editar_janela['listbox_editar'].update(values=[f'PIS {i+1} -> {pis}' for i, pis in enumerate(competencias.get_pis(indice))])

                            elif eventos_edicao == 'todospis_editar' and valores_edicao['todospis_editar']:
                                editar_janela['pis_editar'].update(disabled=True)
                                editar_janela['btn_editar'].update(visible=False)
                                editar_janela['listbox_editar'].update(visible=False)
                                editar_janela['Excluir'].update(visible=False)
                                editar_janela['pis_editar'].update('')
                                
                                apagarpis_editar = True
                                #excluir todos os valores encontrados na listbox
                                
                            elif eventos_edicao == 'umpis_editar' and valores_edicao['umpis_editar']:
                                editar_janela['pis_editar'].update(disabled=False)
                                editar_janela['btn_editar'].update(visible=True)
                                editar_janela['listbox_editar'].update(visible=True)
                                editar_janela['Excluir'].update(visible=True)
                                
                                apagarpis_editar = False
                
                            elif eventos_edicao == 'Excluir':
                                if valores_edicao['listbox_editar']:
                                    pis_lista = competencias.get_pis(indice)
                                    if valores_edicao['listbox_editar'][0][9:] in pis_lista:
                                        indice_pis = competencias.get_pis(indice).index(valores_edicao['listbox_editar'][0][9:])
                                        pis_lista.pop(indice_pis)
                                        competencias.set_pis(indice, pis_lista)
                                        editar_janela['listbox_editar'].update(values=[f'PIS {i+1} -> {pis}' for i, pis in enumerate(competencias.get_pis(indice))])
                                        cont2 -= 1
                                        if cont2 == 0:
                                            editar_janela['btn_editar'].update('+')
                                        else:
                                            editar_janela['btn_editar'].update(cont2)

                            elif eventos_edicao == 'Salvar':
                                todospis = False
                                tipopis = 'UM_PIS'
                                cnpj = valores_edicao['cnpj_editar']
                                pis = valores_edicao['pis_editar']
                                competencia = valores_edicao['competencia_editar']
                                datapag = valores_edicao['datapag_editar']
                                certificado = valores_edicao['certificado_editar']
                                arquivo = valores_edicao['arquivo_editar']
                                if valores_edicao['todospis_editar']:
                                    todospis = True
                                    tipopis = 'TODOS_PIS'

                                erros = []
                                if not cnpj:
                                    erros.append('Campo *CNPJ está vazio.')
                                if not competencia:
                                    erros.append('Campo *Competência está vazio.')
                                if not datapag:
                                    erros.append('Campo *Data pag está vazio.')
                                if not certificado:
                                    erros.append('Campo *Certificado está vazio.')
                                if not arquivo:
                                    erros.append('Campo *Arquivo está vazio.')
                                if erros:
                                    sg.popup_error('\n'.join(erros))
                                    continue
                                erros = []
                                if len(cnpj) < 14:
                                    erros.append(f'Campo *CNPJ não tem dígitos suficientes.')
                                if len(competencia) < 6:
                                    erros.append(f'Campo *Competência não tem dígitos suficientes.')
                                if len(datapag) < 8:
                                    erros.append(f'Campo *Data pag não tem dígitos suficientes.')
                                if erros:
                                    sg.popup_error('\n'.join(erros))
                                    continue
                                erros = []
                                if len(competencias.get_pis(indice)) == 0 and not todospis:
                                    erros.append('Nenhum PIS enviado.')
                                if len(pis) != 0:
                                    erros.append('Campo *PIS tem um valor a ser enviado')
                                if erros:
                                    sg.popup_error('\n'.join(erros))
                                    continue
                                if competencia in competencias.competencia and competencia != valores_popup['opcao'][0]:
                                    sg.popup('Competência já adicionada anteriormente.')
                                    continue
                                
                                if apagarpis_editar == True:
                                    competencias.get_pis(indice).clear()
                                    lista_pis2 = []
                                    editar_janela['listbox_editar'].update(values=lista_pis2)
                                    editar_janela['btn_editar'].update('+')
                                    
                                competencias.set_cnpj(indice, cnpj)
                                competencias.set_competencia(indice, competencia)
                                competencias.set_datapag(indice, datapag)
                                competencias.set_certificado(indice, certificado)
                                competencias.set_arquivo(indice, arquivo)
                                competencias.set_tipopis(indice, tipopis)
                                sg.popup('Alteração salva com sucesso.')
                                historico_janela['opcao'].update(competencias.competencia)
                                break

                            elif eventos_edicao == 'cnpj_editar' or eventos_edicao == 'pis_editar' or eventos_edicao == 'competencia_editar' or eventos_edicao == 'datapag_editar':
                                editar_janela[eventos_edicao].update(value=input_filter(eventos_edicao, valores_edicao[eventos_edicao]))

                        editar_janela.close()
                    else:
                        sg.popup_error('Nenhuma competência selecionada')

                elif eventos_popup == 'Excluir':
                    if valores_popup['opcao']:
                        indice = competencias.competencia.index(valores_popup['opcao'][0])
                        competencias.excluiratributo(indice)
                        historico_janela['opcao'].update(competencias.competencia)
                        sg.popup(f'Competência excluída com sucesso.')
                    else:
                        sg.popup_error('Nenhuma competência selecionada')

            historico_janela.close()

        elif eventos == 'cnpj' or eventos == 'pis' or eventos == 'competencia' or eventos == 'datapag':
            janela[eventos].update(value=input_filter(eventos, valores[eventos]))

    janela.close()

if __name__ == "__main__":
    main()