Attribute VB_Name = "Modulo"
' =========================
' Reais por Extenso
' =========================
' (C)1998 Derlidio Siqueira
' =========================

Option Explicit

Public Function Extenso$(Valor As Variant)

       Static MatrizesCarregadas As Boolean

       Static P1$(1 To 19)
       Static P2$(2 To 9)
       Static P3$(1 To 9)
       
       Const MOEDA_SINGULAR As String = "Real"
       Const MOEDA_PLURAL As String = "Reais"
       Const ESPACO As String = " "
       
       Dim Inteiras$, Decimais$, FocoTrabalho$
       Dim Digito$, Tmp$, Parte$, Bloco$, Frase$, UltimaPalavra$
       Dim Separador%, BlocosAnalisados%, PontoDecimal%
       Dim Posicao%, CurrentStep%

       Dim Plural As Boolean

'      Verifica se o parâmetro recebido é numérico...

       If Not IsNumeric(Valor) Then
          Extenso$ = "Erro: Não é um número!"
          Exit Function
       End If

'      Essa função só consegue lidar com números
'      menores que 1 Trilhão...

       If Valor > 999999999999# Then
          Extenso$ = "Erro: Valor muito grande!"
          Exit Function
       End If
          
'      Se tudo está OK com o parâmetro, prepara-o para
'      análise convertendo-o para o formato string...

       Inteiras$ = Valor & ""

'      Verifica existência de casas decimais...
       
       PontoDecimal% = InStr(Inteiras$, ".") + InStr(Inteiras$, ",")

       If PontoDecimal% Then
          Decimais$ = Mid$(Inteiras$, PontoDecimal% + 1)
          If Decimais$ <> "" Then Decimais$ = Left$(Decimais$ + "00", 2)
          Inteiras$ = Left$(Inteiras$, PontoDecimal% - 1)
       End If

'      Ajusta tamanho dos strings de trabalho para múltiplo de 3...

       If Len(Inteiras$) Mod 3 <> 0 Then Inteiras$ = Space$(3 - (Len(Inteiras$) Mod 3)) + Inteiras$
       If Len(Decimais$) Mod 3 <> 0 Then Decimais$ = Space$(3 - (Len(Decimais$) Mod 3)) + Decimais$
       
       For CurrentStep% = 1 To 2

'          Define ordem de  execução  do  trabalho  de
'          análise: primeiro Inteiras, depois Decimais

           If CurrentStep% = 1 Then
              FocoTrabalho$ = Inteiras$
              Else
              Inteiras$ = Frase$
              Frase$ = ""
              FocoTrabalho$ = Decimais$
           End If

'          A análise do conteúdo de FocoTrabalho$ será
'          feita em blocos de 3 dígitos. Para  isso  é
'          necessário definir um "ponteiro" que  indi-
'          cará o início do bloco a analisar e um con-
'          tador de blocos analisados...

           If FocoTrabalho$ <> "" Then
               
              Separador% = Len(FocoTrabalho$) - 2
              BlocosAnalisados% = 0
              Bloco$ = ""
              
'             Executa processamento do valor enviando os
'             blocos um-a-um p/ a rotina de análise e ar-
'             mazenando o retorno...

              Do
 
                 Tmp$ = Mid$(FocoTrabalho$, Separador, 3)

                 GoSub Extenso_Analise

                 Separador% = Separador% - 3
                 BlocosAnalisados% = BlocosAnalisados% + 1
 
                 If Parte$ <> "" Then

                    Plural = (Trim$(Tmp$) <> "1")
                    
                    Select Case BlocosAnalisados%

                           Case 2

                                If Parte$ = "Um" Then
                                   Parte$ = "Mil": Bloco$ = ""
                                   Else
                                   Bloco$ = "Mil "
                                End If

                           Case 3: If Plural Then Bloco$ = "Milhões " Else Bloco$ = "Milhão "

                           Case 4: If Plural Then Bloco$ = "Bilhões " Else Bloco$ = "Bilhão "

                    End Select
 
                    Parte$ = Parte$ + ESPACO + Bloco$
 
                    If Frase$ <> "" Then
                       If InStr(Frase$, " e ") = 0 Then Parte$ = Parte$ + "e "
                    End If
                    
                    Frase$ = Parte$ + Frase$

                 End If
                       
              Loop Until Separador% < 1
              
           End If
           
       Next CurrentStep%

       If Inteiras$ <> "" Then

          If Inteiras$ <> "Um " Then
             UltimaPalavra$ = Right$(Inteiras$, 6)
             If UltimaPalavra$ = "ilhão " Or UltimaPalavra$ = "lhões " Then Inteiras$ = Inteiras$ + "de "
             Inteiras$ = Inteiras$ + MOEDA_PLURAL + ESPACO
             Else
             Inteiras$ = Inteiras$ + MOEDA_SINGULAR + ESPACO
          End If

          If Frase$ <> "" Then Inteiras$ = Inteiras$ + "e "

       End If

       If Frase$ <> "" Then
          If Frase$ <> "Um " Then
             Frase$ = Frase$ + "Centavos"
             Else
             Frase$ = Frase$ + "Centavo"
          End If
       End If

       Extenso$ = Trim$(Inteiras$ + Frase$)

       Exit Function
  
Extenso_Analise:

'      Verifica se as matrizes de dados já foram
'      carregadas em uma chamada anterior...

       If Not MatrizesCarregadas Then
          
          P1$(1) = "Um"
          P1$(2) = "Dois"
          P1$(3) = "Três"
          P1$(4) = "Quatro"
          P1$(5) = "Cinco"
          P1$(6) = "Seis"
          P1$(7) = "Sete"
          P1$(8) = "Oito"
          P1$(9) = "Nove"
          P1$(10) = "Dez"
          P1$(11) = "Onze"
          P1$(12) = "Doze"
          P1$(13) = "Treze"
          P1$(14) = "Quatorze"
          P1$(15) = "Quinze"
          P1$(16) = "Desesseis"
          P1$(17) = "Dezessete"
          P1$(18) = "Dezoito"
          P1$(19) = "Dezenove"
          
          P2$(2) = "Vinte"
          P2$(3) = "Trinta"
          P2$(4) = "Quarenta"
          P2$(5) = "Cinquenta"
          P2$(6) = "Sessenta"
          P2$(7) = "Setenta"
          P2$(8) = "Oitenta"
          P2$(9) = "Noventa"
          
          P3$(1) = "Cento"
          P3$(2) = "Duzentos"
          P3$(3) = "Trezentos"
          P3$(4) = "Quatrocentos"
          P3$(5) = "Quinhentos"
          P3$(6) = "Seiscentos"
          P3$(7) = "Setecentos"
          P3$(8) = "Oitocentos"
          P3$(9) = "Novecentos"
          
          MatrizesCarregadas = True

       End If

       Parte$ = ""
       
       For Posicao% = 1 To 3

           Digito$ = Mid$(Tmp$, Posicao%, 1)

           If InStr(" 0", Digito$) = 0 Then

              Select Case Posicao%

                     Case 1

                          Parte$ = P3$(Val(Digito$))

                     Case 2

                          If Parte$ <> "" Then Parte$ = Parte$ + " e "

                          If Digito$ = "1" Then
                             Digito$ = Mid$(Tmp$, 2)
                             Parte$ = Parte$ + P1$(Val(Digito$))
                             Exit For
                             Else
                             Parte$ = Parte$ + P2$(Val(Digito$))
                          End If

                     Case 3

                          If Parte$ <> "" Then Parte$ = Parte$ + " e "

                          Parte$ = Parte$ + P1$(Val(Digito$))

              End Select

           End If

       Next Posicao%

       If Parte$ = "Cento" Then Parte$ = "Cem"

       Return

End Function
