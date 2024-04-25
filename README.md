# Conversor de horários da UEM

Esse repositório é um projetinho que eu resolvi fazer para automatizar a importação dos horários que a UEM manda nos começos do semestre para o calendário do Google

O link para o site funcionando é esse daqui: https://murilolcn.github.io/conversor-horarios-uem/

Algumas coisas ainda precisam melhorar pois há ainda algumas limitações, mas até o momento ele funciona bem. O site é bem simples, é simplesmente uma página estática HTML que roda o conversor em JS, nada chique.

Esse projeto é um upgrade gigante em relação a versão antiga do conversor, que eu tinha feito em python em 2022, e - em tese - suporta mais tipos de horários. A UEM utiliza exatamente os mesmos tipos de horários desde os anos 80,
então não há grandes chances desse conversor precisar ser mudado drasticamente ou jogado fora tão cedo... (mas eu não reclamaria se eles atualizassem pra algo bonito ou direto na agenda)

O funcionamento do site é bem simples:

1. Faça o upload do seu horário em PDF - tem que ser o PDF que parece que foi escrito no bloco de notas, com caracteres ASCII.

2. Aperte o botão de "Gerar ICS" e ele irá realizar a conversão do seu horário e baixar um arquivo com a extensão .ICS que pode ser importada na agenda do Google (e em outras também)

3. (opcional) Aperte o botão de "Gerar Excel" para baixar uma planílha simples com o seu horário - inclusive algumas pessoas preferem ter em Excel do que na agenda do Google

4. Importe o ICS gerado na agenda do Google, os passos mais detalhados estão no site
