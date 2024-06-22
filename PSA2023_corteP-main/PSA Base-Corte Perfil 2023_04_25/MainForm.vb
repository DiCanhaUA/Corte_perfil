Imports Mach4
Imports System.IO
Imports System.Threading
Imports System.Windows


Public Class MainForm
    'Variaveis modo escrita'''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim string_texto As String
    Dim file As System.IO.StreamWriter
    Dim counter As Integer
    Dim max_counter As Integer

    '''''''''''''''''''''''''''''''''

    ' Comando Numérico de Máquinas Ferramentas
    ' 10 Nov 2020
    ' ***********************************************
    ' Comentários para organizar correctamente o programa
    ' os nomes dos obejctos devem ter a estrutura:
    '       - tipo objecto + Funçao do Objecto + Opção
    '       - Exemplo:  txt_PosEixoX_Man        
    '                   txt - Text Box
    '                   PosEixoX - Posição eixo X
    '                   Man - Opção modo Manual

    ' DECLARAçÂO DE VARIÁVEIS GLOBAIS
    ' As variáveis Globais são declaradas no módulo: GlobalVars

    ' O CODIGO ASSOCIADO A CADA GRUPO DEVE SER COLOCADO NOS MÓDULOS RESPECTIVOS

    ' QUANDO NECESSÁRIO COLOCAR O CÓDIGO NO MAIN FORM NA ZONA INDICADA.
    ' ver comentários no programa;

    ' Private Mach4.IMach4 _Mach = null;
    ' Private Mach4.IMyScriptObject Script = null;

    Public mach As Mach4.IMach4
    Public scriptObject As Mach4.IMyScriptObject

    ' Variável de demonstração do ficheiro de configuração
    Public Manual_ManualFeedRate As Integer


    Private Sub MainForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim aux As String

        ' Propriedades Iniciais
        ' Modo Manual
        StatusStripLbl_Modo.Text = "Modo Manual"

        ' Iniciar comunicações com o Match3
        mach = GetObject(, "Mach4.Document")
        scriptObject = mach.GetScriptDispatch()

        ' Necessário confirmar se o Match3 está OK

        ' Definições do temporizador para leitura de posição dos eixos (Match3)
        tmr_match3.Interval = 250
        tmr_match3.Enabled = True

        ' Ficheiro de configurações - Modo Manual
        FileOpen(1, "Config\CmdNum_Manual.Ini", OpenMode.Input)
        For i = 1 To 8
            aux = LineInput(1)
        Next
        ' Leitura da velocidade de deslocamento
        aux = LineInput(1)
        Dim words As String() = aux.Split(New Char() {" "c})

        Manual_ManualFeedRate = Integer.Parse(words(2))

        ' Visualizaçºao da Velocidade de Deslocamento
        txt_ManFeedRate.Text = Manual_ManualFeedRate.ToString

        ' PrintLine  para escerever em ficheiro...
        'PrintLine (1, "teste teste")

        FileClose(1)

        pic_gifAut.Visible = False
        pic_gifMan.Visible = False


    End Sub

    Private Sub TabCtrl_Option_Selecting(sender As Object, e As TabControlCancelEventArgs) Handles TabCtrl_Option.Selecting
        ' Selecao de uma determinada página
        PaginaSelecionada = TabCtrl_Option.SelectedIndex.ToString

        Select Case (PaginaSelecionada)
            Case "0"
                ' Modo Manual
                StatusStripLbl_Modo.Text = "Modo Manual"

            Case "1"
                ' Modo Automático
                StatusStripLbl_Modo.Text = "Modo Automático"
            Case "2"
                ' Programas
                StatusStripLbl_Modo.Text = "Programas G Code"
            Case "3"
                ' Tabelas
                StatusStripLbl_Modo.Text = "Tabelas"
            Case "4"
                ' Parametros
                StatusStripLbl_Modo.Text = "Parametros"
            Case "5"
                ' Avisos / Help
                StatusStripLbl_Modo.Text = "Avisos"
            Case "6"
                ' Modo escrita
                StatusStripLbl_Modo.Text = "ABC"
            Case Else
                ' Outros....
                StatusStripLbl_Modo.Text = "Outros... Erro..."

        End Select


    End Sub

    '*******************************************************************
    '*******************************************************************


    ' *******************************************
    ' ROTINAS MODO MANUAL
    ' *******************************************

    Private Sub btn_ManSet_Click(sender As Object, e As EventArgs) Handles btn_ManSet.Click
        ' Zero Maquina
        ' A Borges - 28 Nov 2020
        Manual_SetRefPoint()
        pic_gifMan.Visible = True
    End Sub

    Private Sub btn_ManXmenos_MouseDown(sender As Object, e As MouseEventArgs) Handles btn_ManXmenos.MouseDown
        ' Deslocamento Eixo X-
        scriptObject.DoOEMButton(308)
        pic_gifMan.Visible = True
    End Sub

    Private Sub btn_ManXmenos_MouseUp(sender As Object, e As MouseEventArgs) Handles btn_ManXmenos.MouseUp
        ' Quando Liberta o X- Interrompe movimento
        scriptObject.DoOEMButton(334)
        pic_gifMan.Visible = False
    End Sub

    Private Sub btn_ManXmais_MouseUp(sender As Object, e As MouseEventArgs) Handles btn_ManXmais.MouseUp
        ' Quando Liberta o X+ Interrompe movimento
        scriptObject.DoOEMButton(334)
        pic_gifMan.Visible = False
    End Sub

    Private Sub btn_ManXmais_MouseDown(sender As Object, e As MouseEventArgs) Handles btn_ManXmais.MouseDown
        ' Movimento Manual X+
        scriptObject.DoOEMButton(307)
        pic_gifMan.Visible = True
    End Sub

    Private Sub btn_ManMovToPos_Click(sender As Object, e As EventArgs) Handles btn_ManMovToPos.Click
        ' Modo Manual
        ' Move para o ponto definido...
        Dim auxMDI As String

        ' Abertura das comunicações com o Match3
        mach = GetObject(, "Mach4.Document")
        scriptObject = mach.GetScriptDispatch()

        auxMDI = "G90 G01 F" + Manual_ManualFeedRate.ToString + " X" + txt_PosManEixoX.Text

        scriptObject.Code(auxMDI.Trim())

        pic_gifMan.Visible = True


    End Sub

    ' *******************************************
    ' ROTINAS MODO AUTOMÁTICO
    ' *******************************************



    ' *******************************************
    ' ROTINAS PROGRAMA
    ' *******************************************


    ' *******************************************
    ' ROTINAS TABELAS
    ' *******************************************




    ' *******************************************
    ' ROTINAS PARAMETROS
    ' *******************************************



    ' *******************************************
    ' ROTINAS AVISOS
    ' *******************************************

    Private Sub Button_close_Click(sender As Object, e As EventArgs) Handles Button_close.Click
        ' Terminar programa
        End
    End Sub


    Private Sub tmr_match3_Tick(sender As Object, e As EventArgs) Handles tmr_match3.Tick
        ' Visualização coordenadas correntes do eixo X
        txt_ManPosX.Text = Format(scriptObject.GetOEMDRO(800), "###,##0.#0")
        txt_PosAutEixoX.Text = Format(scriptObject.GetOEMDRO(800), "###,##0.#0")

    End Sub

    Private Sub trackBar_AutFeed_MouseUp(sender As Object, e As MouseEventArgs) Handles trackBar_ManFeed.MouseUp
        txt_ManFeedRate.Text = trackBar_ManFeed.Value.ToString

    End Sub



    Function Cmd_Line(ByVal posX As String, ByVal velocidade As String) As String
        ' Execução cmd MDI
        Dim cmd As String = ""
        cmd = " G91" + " G01 X" + posX + " F" + velocidade
        Cmd_Line = cmd
    End Function





    Private Sub btn_ManValidCodeG_Click(sender As Object, e As EventArgs) Handles btn_ManValidCodeG.Click
        ' Ver rotina de modo automático 
        ' muito semelhante para se poder optimizar
        ' 

        Dim VelocidadeEixos As Double
        VelocidadeEixos = 0

        ' MDI

        Dim filename As String
        Dim linha As String
        txt_ManCodG.Text = ""


        ' Verifica se existe o ficheiro
        filename = "C:\tmp\PSA2023.tap"
        'txt_Prg_FileName.Text = filename

        ' Ficheiro de configurações - Modo Manual
        FileOpen(1, filename, OpenMode.Output)

        ' PrintLine  para escerever em ficheiro...
        'PrintLine (1, "teste teste")

        ' Write single line to New file.
        linha = "; (PSA2023)"
        PrintLine(1, linha)
        txt_ManCodG.Text += linha + vbCrLf

        linha = "N00 M27"        'primeiro corte/21/05
        PrintLine(1, linha)
        txt_ManCodG.Text += linha + vbCrLf

        'linha = "N01 M4" 'desativar cerra
        'PrintLine(1, linha)
        'txt_ManCodG.Text += linha + vbCrLf



        linha = "N10 G4 P2" 'tempo de descanso
        PrintLine(1, linha)
        txt_ManCodG.Text += linha + vbCrLf


        VelocidadeEixos = Double.Parse(txt_VelX_1.Text)
        linha = Cmd_Line(txt_PosX_1.Text.ToString(), VelocidadeEixos.ToString())

        PrintLine(1, "N20" + linha)



        txt_ManCodG.Text += "N20 " + linha + vbCrLf ' Visualização do programa




        '---=== M3 Macro ===---                       'Não no VB script window??

        'ActivateSignal(Output3)   'Release the break

        'Sleep(200)                'Wait 200mS

        'While Not isActive(Input4)    'Sensor da cerra baixo

        '    sleep(200)

        'End While

        ''DoSpinCW()               'Start spindle

        'sleep(1000)

        'DeActivateSignal(Output3) 'Apply the break

        'Sleep(200)                'Wait 200mS

        'While Not isActive(Input3)                        'Não if??

        '    sleep(200)

        'End While
        linha = "N30 M27" 'ativar corte completo 21/05
        PrintLine(1, linha)
        txt_ManCodG.Text += linha + vbCrLf

        ' linha = "N20 M10" 'ativar eletrovalvulas OUTPUT 7
        ' PrintLine(1, linha)
        ' txt_ManCodG.Text += linha + vbCrLf

        ' linha = "N30 M3" 'ativar cerra   OUTPUT 5 ABRIR OUTPUT 6
        'PrintLine(1, linha)
        ' txt_ManCodG.Text += linha + vbCrLf

        'linha = "N40 M4 " 'desativar cerra
        ' PrintLine(1, linha)
        ' txt_ManCodG.Text += linha + vbCrLf


        ' linha = "N50 M11" 'desativar eletrovalvulas  OUTPUT 8
        ' PrintLine(1, linha)
        ' txt_ManCodG.Text += linha + vbCrLf


        linha = "N40 G4 P2" 'ativar corte completo 21/05
        PrintLine(1, linha)
        txt_ManCodG.Text += linha + vbCrLf



        VelocidadeEixos = Double.Parse(txt_VelX_2.Text)
        linha = Cmd_Line(txt_PosX_2.Text.ToString(), VelocidadeEixos.ToString())

        PrintLine(1, "N50" + linha)



        txt_ManCodG.Text += "N50 " + linha + vbCrLf ' Visualização do programa

        linha = "N60 M27" 'ativar corte completo 21/05
        PrintLine(1, linha)
        txt_ManCodG.Text += linha + vbCrLf

        ' linha = "N90 M10" 'ativar eletrovalvulas
        ' PrintLine(1, linha)
        'txt_ManCodG.Text += linha + vbCrLf

        ' linha = "N100 M3" 'ativar cerra
        ' PrintLine(1, linha)
        'txt_ManCodG.Text += linha + vbCrLf

        ' linha = "N110 M11" 'desativar eletrovalvulas
        ' PrintLine(1, linha)
        ' txt_ManCodG.Text += linha + vbCrLf


        linha = "N70 G4 P2" 'ativar corte completo 21/05
        PrintLine(1, linha)
        txt_ManCodG.Text += linha + vbCrLf

        'linha = "N100 M7 " 'ativar eletro valvulas
        'PrintLine(1, linha)
        'txt_ManCodG.Text += linha + vbCrLf

        'linha = "N110 M3" 'ativar cerra
        'PrintLine(1, linha)
        'txt_ManCodG.Text += linha + vbCrLf

        'linha = "N120 M4 " 'desativar cerra
        'PrintLine(1, linha)
        'txt_ManCodG.Text += linha + vbCrLf

        'linha = "N130 M8 " 'desativar eletro valvulas
        'PrintLine(1, linha)
        'txt_ManCodG.Text += linha + vbCrLf


        VelocidadeEixos = Double.Parse(txt_VelX_3.Text)
        linha = Cmd_Line(txt_PosX_3.Text.ToString(), VelocidadeEixos.ToString())

        PrintLine(1, "N80" + linha)



        txt_ManCodG.Text += "N80 " + linha + vbCrLf ' Visualização do programa

        linha = "N90 M27" 'ativar corte completo 21/05
        PrintLine(1, linha)
        txt_ManCodG.Text += linha + vbCrLf




        'linha = "N160 M10" 'ativar eletrovalvulas
        ' PrintLine(1, linha)
        'txt_ManCodG.Text += linha + vbCrLf

        ' linha = "N170 M3" 'ativar cerra
        ' PrintLine(1, linha)
        'txt_ManCodG.Text += linha + vbCrLf

        ' linha = "N180 M11" 'desativar eletrovalvulas
        ' PrintLine(1, linha)
        'txt_ManCodG.Text += linha + vbCrLf

        linha = "N100 G4 P2" 'ativar corte completo 21/05
        PrintLine(1, linha)
        txt_ManCodG.Text += linha + vbCrLf
        'linha = "N170 M7 " 'ativar eletro valvulas
        'PrintLine(1, linha)
        'txt_ManCodG.Text += linha + vbCrLf

        'linha = "N180 M3" 'ativar cerra
        'PrintLine(1, linha)
        'txt_ManCodG.Text += linha + vbCrLf

        'linha = "N190 M4 " 'desativar cerra
        'PrintLine(1, linha)
        'txt_ManCodG.Text += linha + vbCrLf

        'linha = "N200 M8 " 'desativar eletro valvulas
        'PrintLine(1, linha)
        'txt_ManCodG.Text += linha + vbCrLf

        'linha = "N210 M47 " 'codigo do professor
        'PrintLine(1, linha)
        'txt_ManCodG.Text += linha + vbCrLf

        VelocidadeEixos = Double.Parse(txt_VelX_4.Text)
        linha = Cmd_Line(txt_PosX_4.Text.ToString(), VelocidadeEixos.ToString())

        PrintLine(1, "N110" + linha)



        txt_ManCodG.Text += "N110 " + linha + vbCrLf ' Visualização do programa

        linha = "N120 M27" 'ativar corte completo 21/05
        PrintLine(1, linha)
        txt_ManCodG.Text += linha + vbCrLf


        'linha = "N230 M10" 'ativar eletrovalvulas
        ' PrintLine(1, linha)
        ' txt_ManCodG.Text += linha + vbCrLf

        linha = "N130 G4 P2" 'ativar corte completo 21/05
        PrintLine(1, linha)
        txt_ManCodG.Text += linha + vbCrLf

        ' linha = "N240 M3" 'ativar cerra
        ' PrintLine(1, linha)
        'txt_ManCodG.Text += linha + vbCrLf

        ' linha = "N250 M11" 'desativar eletrovalvulas
        ' PrintLine(1, linha)
        ' txt_ManCodG.Text += linha + vbCrLf


        VelocidadeEixos = Double.Parse(txt_VelX_5.Text)
        linha = Cmd_Line(txt_PosX_5.Text.ToString(), VelocidadeEixos.ToString())

        PrintLine(1, "N140" + linha)



        txt_ManCodG.Text += "N140 " + linha + vbCrLf ' Visualização do programa

        linha = "N150 M27" 'ativar corte completo 21/05
        PrintLine(1, linha)
        txt_ManCodG.Text += linha + vbCrLf

        ' linha = "N300 M10" 'ativar eletrovalvulas
        'PrintLine(1, linha)
        'txt_ManCodG.Text += linha + vbCrLf

        linha = "N160 G4 P2" 'ativar corte completo 21/05
        PrintLine(1, linha)
        txt_ManCodG.Text += linha + vbCrLf

        'linha = "N310 M3" 'ativar cerra
        ' PrintLine(1, linha)
        ' txt_ManCodG.Text += linha + vbCrLf

        'linha = "N320 M11" 'desativar eletrovalvulas
        'PrintLine(1, linha)
        ' txt_ManCodG.Text += linha + vbCrLf

        '///////////////////////////////'

        VelocidadeEixos = Double.Parse(txt_VelX_6.Text)
        linha = Cmd_Line(txt_PosX_6.Text.ToString(), VelocidadeEixos.ToString())

        PrintLine(1, "N170" + linha)



        txt_ManCodG.Text += "N170 " + linha + vbCrLf ' Visualização do programa

        linha = "N180 M27" 'ativar corte completo 21/05
        PrintLine(1, linha)
        txt_ManCodG.Text += linha + vbCrLf

        ' linha = "N370 M10" 'ativar eletrovalvulas
        ' PrintLine(1, linha)
        ' txt_ManCodG.Text += linha + vbCrLf

        linha = "N190 G4 P2" 'ativar corte completo 21/05
        PrintLine(1, linha)
        txt_ManCodG.Text += linha + vbCrLf

        ' linha = "N380 M3" 'ativar cerra
        '  PrintLine(1, linha)
        ' txt_ManCodG.Text += linha + vbCrLf

        'linha = "N390 M11" 'desativar eletrovalvulas
        'PrintLine(1, linha)
        ' txt_ManCodG.Text += linha + vbCrLf

        '////////////////////////////////////////'

        VelocidadeEixos = Double.Parse(txt_VelX_7.Text)
        linha = Cmd_Line(txt_PosX_7.Text.ToString(), VelocidadeEixos.ToString())

        PrintLine(1, "N200" + linha)



        txt_ManCodG.Text += "N200 " + linha + vbCrLf ' Visualização do programa

        linha = "N210 M27" 'ativar corte completo 21/05
        PrintLine(1, linha)
        txt_ManCodG.Text += linha + vbCrLf

        'linha = "N440 M10" 'ativar eletrovalvulas
        'PrintLine(1, linha)
        'txt_ManCodG.Text += linha + vbCrLf

        'linha = "N450 M3" 'ativar cerra
        'PrintLine(1, linha)
        'txt_ManCodG.Text += linha + vbCrLf

        ' linha = "N460 M11" 'desativar eletrovalvulas
        'PrintLine(1, linha)
        ' txt_ManCodG.Text += linha + vbCrLf

        '//////////////////////////'

        VelocidadeEixos = Double.Parse(txt_VelX_8.Text)
        linha = Cmd_Line(txt_PosX_8.Text.ToString(), VelocidadeEixos.ToString())

        PrintLine(1, "N220" + linha)



        txt_ManCodG.Text += "N220 " + linha + vbCrLf ' Visualização do programa

        linha = "N230 M27" 'ativar corte completo 21/05
        PrintLine(1, linha)
        txt_ManCodG.Text += linha + vbCrLf

        ' linha = "N510 M10" 'ativar eletrovalvulas
        ' PrintLine(1, linha)
        'txt_ManCodG.Text += linha + vbCrLf

        linha = "N240 G4 P2000" 'ativar corte completo 21/05
        PrintLine(1, linha)
        txt_ManCodG.Text += linha + vbCrLf
        ' linha = "N520 M3" 'ativar cerra
        ' PrintLine(1, linha)
        ' txt_ManCodG.Text += linha + vbCrLf

        ' linha = "N530 M11" 'desativar eletrovalvulas
        ' PrintLine(1, linha)
        'txt_ManCodG.Text += linha + vbCrLf

        '//////////////////////////'

        VelocidadeEixos = Double.Parse(txt_VelX_9.Text)
        linha = Cmd_Line(txt_PosX_9.Text.ToString(), VelocidadeEixos.ToString())

        PrintLine(1, "N250" + linha)



        txt_ManCodG.Text += "N250 " + linha + vbCrLf ' Visualização do programa

        linha = "N260 M27" 'ativar corte completo 21/05
        PrintLine(1, linha)
        txt_ManCodG.Text += linha + vbCrLf

        'linha = "N580 M10" 'ativar eletrovalvulas
        ' PrintLine(1, linha)
        ' txt_ManCodG.Text += linha + vbCrLf

        linha = "N270 G4 P2" 'ativar corte completo 21/05
        PrintLine(1, linha)
        txt_ManCodG.Text += linha + vbCrLf

        ' linha = "N590 M3" 'ativar cerra
        ' PrintLine(1, linha)
        ' txt_ManCodG.Text += linha + vbCrLf

        'linha = "N600 M11" 'desativar eletrovalvulas
        ' PrintLine(1, linha)
        ' txt_ManCodG.Text += linha + vbCrLf

        '//////////////////////////'

        VelocidadeEixos = Double.Parse(txt_VelX_10.Text)
        linha = Cmd_Line(txt_PosX_10.Text.ToString(), VelocidadeEixos.ToString())

        PrintLine(1, "N280" + linha)



        txt_ManCodG.Text += "N280 " + linha + vbCrLf ' Visualização do programa

        linha = "N290 M27" 'ativar corte completo 21/05
        PrintLine(1, linha)
        txt_ManCodG.Text += linha + vbCrLf

        'linha = "N640 M10" 'ativar eletrovalvulas
        'PrintLine(1, linha)
        'txt_ManCodG.Text += linha + vbCrLf

        linha = "N300 G4 P2" 'ativar corte completo 21/05
        PrintLine(1, linha)
        txt_ManCodG.Text += linha + vbCrLf

        ' linha = "N650 M3" 'ativar cerra
        ' PrintLine(1, linha)
        'txt_ManCodG.Text += linha + vbCrLf

        ' linha = "N660 M11" 'desativar eletrovalvulas
        ' PrintLine(1, linha)
        ' txt_ManCodG.Text += linha + vbCrLf


        linha = "N310  M30" 'fim do programa
        PrintLine(1, linha)
        txt_ManCodG.Text += linha + vbCrLf

        FileClose(1)

    End Sub

    Private Sub btn_ManRun_Click(sender As Object, e As EventArgs) Handles btn_ManRun.Click
        ' Verifica se existe o ficheiro
        Dim filename As String = "C:tmp\PSA2023.tap"
        ' txt_Prg_FileName.Text = filename;

        scriptObject.LoadFile(filename)
        Thread.Sleep(1000)
        scriptObject.RunFile()
        pic_gifAut.Visible = True

    End Sub

End Class
