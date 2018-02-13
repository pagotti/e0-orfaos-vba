Office.initialize = function (reason) {
    function jExtenso(valor) {
        var matrizesCarregadas = false;
        var p1 = [];
        var p2 = [];
        var p3 = [];
        var moedaSingular = "Real";
        var moedaPlural = "Reais";
        var espaco = " ";
        var inteiras = "", decimais = "", focoTrabalho = "";
        var digito = "", tmp = "", parte = "", bloco = "", frase = "", ultimaPalavra = "";
        var separador = 0, blocosAnalisados = 0, pontoDecimal = 0;
        var posicao = 0, currentStep = 0;
        var plural = false;
        // Verifica se o parâmetro recebido é numérico...
        if (isNaN(parseFloat(valor))) {
            return "Erro: Não é um número!";
        }
        // Essa função só consegue lidar com números
        // menores que 1 Trilhão...
        if (parseFloat(valor) > 999999999999) {
            return "Erro: Valor muito grande!";
        }
        // Se tudo está OK com o parâmetro, prepara-o para
        // análise convertendo-o para o formato string ...
        inteiras = parseFloat(valor) + "";
        // Verifica existência de casas decimais ...
        pontoDecimal = inteiras.indexOf(".");
        if (pontoDecimal != -1) {
            // Diferente do VBA, a primeira posição de uma string é 0 (zero).
            decimais = inteiras.substring(pontoDecimal + 1);
            if (decimais != "") {
                decimais = (decimais + "00").substring(0, 2);
            }
            inteiras = inteiras.substring(0, pontoDecimal);
        }
        // Ajusta tamanho dos strings de trabalho para múltiplo de 3...
        if (inteiras.length % 3 != 0) {
            inteiras = ("   ").substr(inteiras.length % 3) + inteiras;
        }
        if (decimais.length % 3 != 0) {
            decimais = ("   ").substr(decimais.length % 3) + decimais;
        }
        for (currentStep = 1; currentStep < 3; currentStep++) {
            // Define ordem de  execução  do  trabalho  de
            // análise: primeiro Inteiras, depois Decimais
            if (currentStep == 1) {
                focoTrabalho = inteiras;
            }
            else {
                inteiras = frase;
                frase = "";
                focoTrabalho = decimais;
            }
            //  A análise do conteúdo de FocoTrabalho será
            //  feita em blocos de 3 dígitos. Para  isso  é
            //  necessário definir um "ponteiro" que  indi-
            //  cará o início do bloco a analisar e um con-
            //  tador de blocos analisados ... 
            if (focoTrabalho != "") {
                separador = focoTrabalho.length - 3;
                blocosAnalisados = 0;
                bloco = "";
                // Executa processamento do valor enviando os
                // blocos um-a-um p/ a rotina de análise e ar-
                // mazenando o retorno ..
                do {
                    tmp = focoTrabalho.substr(separador, 3);
                    extensoAnalise();
                    separador -= 3;
                    blocosAnalisados += 1;
                    if (parte != "") {
                        plural = tmp.trim() != "1";
                        switch (blocosAnalisados) {
                            case 2:
                                if (parte == "Um") {
                                    parte = "Mil";
                                    bloco = "";
                                }
                                else {
                                    bloco = "Mil ";
                                }
                                break;
                            case 3:
                                if (plural) {
                                    bloco = "Milhões ";
                                }
                                else {
                                    bloco = "Milhão ";
                                }
                                break;
                            case 4:
                                if (plural) {
                                    bloco = "Bilhões ";
                                }
                                else {
                                    bloco = "Bilhão ";
                                }
                                break;
                        }
                        parte = parte + espaco + bloco;
                        if (frase != "") {
                            if (frase.indexOf(" e ") == -1) {
                                parte = parte + "e ";
                            }
                        }
                        frase = parte + frase;
                    }
                } while (separador >= 0);
            }
        }
        if (inteiras != "") {
            if (inteiras != "Um ") {
                ultimaPalavra = inteiras.substring(inteiras.length - 6);
                if (ultimaPalavra == "ilhão " || ultimaPalavra == "lhões ") {
                    inteiras = inteiras + "de ";
                }
                inteiras = inteiras + moedaPlural + espaco;
            }
            else {
                inteiras = inteiras + moedaSingular + espaco;
            }
            if (frase != "") {
                inteiras = inteiras + "e ";
            }
        }
        if (frase != "") {
            if (frase != "Um ") {
                frase = frase + "Centavos";
            }
            else {
                frase = frase + "Centavo";
            }
        }
        return (inteiras + frase).trim() + "";
        function extensoAnalise() {
            //  Verifica se as matrizes de dados já foram
            //  carregadas em uma chamada anterior ...
            if (!matrizesCarregadas) {
                p1[1] = "Um";
                p1[2] = "Dois";
                p1[3] = "Três";
                p1[4] = "Quatro";
                p1[5] = "Cinco";
                p1[6] = "Seis";
                p1[7] = "Sete";
                p1[8] = "Oito";
                p1[9] = "Nove";
                p1[10] = "Dez";
                p1[11] = "Onze";
                p1[12] = "Doze";
                p1[13] = "Treze";
                p1[14] = "Quatorze";
                p1[15] = "Quinze";
                p1[16] = "Desesseis";
                p1[17] = "Dezessete";
                p1[18] = "Dezoito";
                p1[19] = "Dezenove";
                p2[2] = "Vinte";
                p2[3] = "Trinta";
                p2[4] = "Quarenta";
                p2[5] = "Cinquenta";
                p2[6] = "Sessenta";
                p2[7] = "Setenta";
                p2[8] = "Oitenta";
                p2[9] = "Noventa";
                p3[1] = "Cento";
                p3[2] = "Duzentos";
                p3[3] = "Trezentos";
                p3[4] = "Quatrocentos";
                p3[5] = "Quinhentos";
                p3[6] = "Seiscentos";
                p3[7] = "Setecentos";
                p3[8] = "Oitocentos";
                p3[9] = "Novecentos";
                matrizesCarregadas = true;
            }
            parte = "";
            var ExitFor = false;
            for (posicao = 0; posicao < 3; posicao++) {
                digito = tmp.substr(posicao, 1);
                if ((" 0").indexOf(digito) == -1) {
                    switch (posicao) {
                        case 0:
                            parte = p3[parseInt(digito)];
                            break;
                        case 1:
                            if (parte != "") {
                                parte = parte + " e ";
                            }
                            if (digito == "1") {
                                digito = tmp.substring(1);
                                parte = parte + p1[parseInt(digito)];
                                ExitFor = true;
                            }
                            else {
                                parte = parte + p2[parseInt(digito)];
                            }
                            break;
                        case 2:
                            if (parte != "") {
                                parte = parte + " e ";
                            }
                            parte = parte + p1[parseInt(digito)];
                            break;
                    }
                    if (ExitFor)
                        break;
                }
            }
            if (parte == "Cento") {
                parte = "Cem";
            }
            return;
        }
    }
    // Monta a definição da função que será utilizada no Excel:
    Excel.Script.CustomFunctions = {};
    Excel.Script.CustomFunctions["EZERO"] = {};
    Excel.Script.CustomFunctions["EZERO"]["JEXTENSO"] =
        {
            call: jExtenso,
            description: "Retorna valor em reais por extenso.",
            helpUrl: "https://localhost/function/help.html",
            result: {
                resultType: Excel.CustomFunctionValueType.string,
                resultDimensionality: Excel.CustomFunctionDimensionality.scalar
            },
            parameters: [
                {
                    name: "valor",
                    description: "valor em reais",
                    valueType: Excel.CustomFunctionValueType.number,
                    valueDimensionality: Excel.CustomFunctionDimensionality.scalar
                }
            ],
            options: { batch: false, stream: false }
        };
    // Adiciona a função ao Excel:
    Excel.run(function (context) {
        context.workbook.customFunctions.addAll();
        return context.sync();
    })["catch"](function (error) { });
};
