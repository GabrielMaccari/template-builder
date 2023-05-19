<div align="center"/>
  <p>
    <h1>Template Builder</h1>
      <h4>
        Ferramenta para preenchimento semiautomático dos cabeçalhos da caderneta de campo compilada para a disciplina de Mapeamento Geológico do curso de graduação
        em Geologia da UFSC.
      </h4>
      </br>
  <p>
</div> 

![qt_windows](https://github.com/FrostPredator/template-builder/assets/114439033/b93eee2c-6dac-4a13-a420-60df72015317)

## Qual versão baixar?
- **Windows:** Ambas as versões (Qt e wx) funcionam perfeitamente. Baixe a versão Qt se você usa o tema escuro do Windows 11.
- **Linux:** Baixe a versão wx, pois ela se adaptará ao tema do seu Desktop Environment (testado no GNOME e Cinnamon). A versão Qt pode não funcionar em distros que não utilizem Wayland.

## Instruções de uso
### Como preencher a tabela de pontos
#### Passo 1: Configurando as listas de opções
- Abra o arquivo “tabela_caderneta.xlsx” utilizando o Microsoft Office Excel ou outro software de edição de planilhas.
- Acesse a segunda aba do arquivo (planilha “Listas”).

![aba_listas](https://github.com/FrostPredator/template-builder/assets/114439033/311d51e3-5be9-40ae-a907-10797058d486)

- Nos campos delimitados, preencha as **unidades** (suítes, complexos, grupos, etc.), **unidades litoestratigráficas** (formações, granitos, unidades, etc.) e **estruturas** observadas em campo. **Ao preencher a lista de estruturas, em específico, não use espaços, acentos e cedilha. Insira o nome da estrutura, seguido pela sigla dela entre parênteses**, conforme os exemplos que aparecem abaixo da lista. **Não recorte células na lista de estruturas**, pois isso desconfigurará os cabeçalhos das colunas de estruturas na planilha Geral.

![preenchimento_listas](https://github.com/FrostPredator/template-builder/assets/114439033/0704e83e-39bc-461d-a248-388d32692220)

<sub>Obs: A sigla da estrutura deve estar entre parênteses para que a ferramenta de preenchimento do template da caderneta possa detectá-la. Caso não haja sigla, o nome completo inserido será utilizado.</sub>

As listas de unidades aparecerão depois como opções para o preenchimento desses campos na aba Geral:

![image](https://github.com/FrostPredator/template-builder/assets/114439033/af462b47-e22f-4f01-b565-ce70b106f32b)

O preenchimento da lista de estruturas, por outro lado, adiciona as estruturas inseridas como colunas na primeira aba:

![image](https://github.com/FrostPredator/template-builder/assets/114439033/3d045c8f-f772-4209-8337-86465ab0aad9)

#### Passo 2: Preenchendo a tabela pós-campo
A tabela deve ser preenchida diretamente a partir dos dados da caderneta de campo.
Evite colar dados de outras tabelas e softwares. **Quando colar quaisquer dados, utilize a colagem apenas de valores**. No Microsoft Office Excel, essa opção pode ser encontrada clicando com o botão direito do mouse na célula alvo da colagem, acessando a opção “Colar Especial...” e selecionando “Valores” (símbolo de prancheta com “123”). Isso impede que a formatação de validação dos dados seja substituída.
Você pode utilizar acentos, cedilhas e caracteres especiais no preenchimento dos campos.
**NÃO insira ou exclua colunas.**
**NÃO edite os nomes das colunas.**
**NÃO troque a ordem das colunas.**
**Siga as instruções de preenchimento abaixo para cada coluna.**
**Salve a tabela usando os formatos .xlsx ou .xlsm.**

- **Ponto:** O código do ponto de campo. Ex: PTI-2001. **Não deixe em branco**. Preencha na ordem de numeração.
- **Disciplina:** A disciplina na qual o ponto foi visitado pela primeira vez. Preencha com “Mapeamento Geológico I” ou “Mapeamento Geológico II”. **Não deixe em branco**. Preencha apenas de forma contínua (depois que preencher uma linha com “Mapeamento Geológico II”, não preencha nenhuma linha seguinte com “Mapeamento Geológico I”).
- **SRC:** O sistema de referência de coordenadas configurado no GPS utilizado em campo. Ex: "WGS 84 / UTM zone 22S". **Não deixe em branco**.
- **Easting:** A coordenada UTM leste (easting) do ponto, em metros. Insira apenas números. **Não deixe em branco**.
- **Northing:** A coordenada UTM norte (northing) do ponto, em metros. Insira apenas números. **Não deixe em branco**.
- **Altitude:** A altitude do ponto, em metros. Insira apenas números.
- **Toponimia:** A toponímia do local ou um local de referência próximo ao ponto. Utilize apenas referências duradouras e que possam ajudar alguém que nunca esteve no local a encontrar o ponto no futuro.
- **Data:** A data de visita ao ponto, no formato dia/mês/ano. Ex: "01/08/1997". **Não deixe em branco**.
- **Equipe:** Os nomes dos integrantes da equipe que visitou o ponto, incluindo professores, separados por vírgula e espaço. Utilize apenas o último sobrenome de cada integrante, e nenhum nome do meio. Ex: "Ana Sutili, Gabriel Maccari, Vicente Wetter, Luana Florisbal". **Não deixe em branco**.
- **Ponto_de_controle:** Se o ponto em questão é apenas um ponto de controle, ou se possui afloramento. Preencha com “Sim” ou “Não” (sem aspas, com acento, inicial maiúscula). **Não deixe em branco**.
- **Numero_de_amostras:** O número de amostras coletadas no ponto. Preencha apenas com números inteiros. Preencha com zero caso nenhuma amostra tenha sido coletada. **Não deixe em branco**.
- **Possui_croquis:** Se foram feitos croquis para ilustrar alguma feição no ponto (e se eles serão incluídos na caderneta). Preencha com “Sim” ou “Não” (sem aspas, com acento, inicial maiúscula). **Não deixe em branco**.
- **Possui_fotos:** Se foram tiradas fotos do ponto (e se elas serão incluídas na caderneta). Preencha com “Sim” ou “Não” (sem aspas, com acento, inicial maiúscula). **Não deixe em branco**.

<sub>Obs: Os campos a seguir (Tipo_de_afloramento, In_situ, Grau_de_intemperismo, Unidade, Unidade_litoestratigrafica e campos de medidas estruturais) devem ser preenchidos apenas nos pontos que contêm afloramento, e devem ser deixados em branco nos pontos de controle.</sub>

- **Tipo_de_afloramento:** O tipo de afloramento presente no ponto em questão. Ex: "Corte de estrada", "Barranco", "Drenagem", etc.
- **In_situ:** Se as rochas descritas no ponto encontravam-se in situ ou se foram transportadas de outro local (como no caso de matacões rolados morro abaixo ou seixos em uma drenagem). Preencha com “Sim” ou “Não” (sem aspas, com acento, inicial maiúscula).
- **Grau_de_intemperismo:** O grau de alteração do afloramento frente às intempéries. Preencha com “Baixo”, “Médio” ou “Alto” (sem aspas, com acento, inicial maiúscula).
- **Unidade:** A unidade maior na qual a litologia principal do ponto está contida. O preenchimento deste campo deve ser feito conforme as unidades listadas na segunda aba da planilha. Ex: "Complexo Metamórfico Brusque", "Suíte Valsungana", "Coberturas Cenozoicas", etc. Caso o ponto em questão seja um ponto de contato entre duas unidades, acrescente na aba de Listas uma unidade mista, separada por “/” (Ex: "Grupo Itararé / Grupo Itajaí"), e então preencha o ponto com a unidade adicionada.
- **Unidade_litoestratigrafica:** A unidade litoestratigráfica específica na qual a litologia principal do ponto está contida. O preenchimento deste campo deve ser feito conforme as unidades litoestratigráficas listadas na segunda aba da planilha. Ex: "Formação Rio Bonito", "Granodiorito Estaleiro", etc. Caso o ponto em questão seja um ponto de contato entre duas unidades, acrescente na aba de Listas uma unidade mista, separada por “/” (Ex: "Formação Taciba / Formação Campo Mourão"), e então preencha o ponto com a unidade adicionada.
- **_Campos de estruturas_:** Preencha com as medidas tiradas para a estrutura em questão, separadas por vírgula e espaço. Caso haja mais de uma medida da mesma estrutura no mesmo ponto, separe-as com vírgula e espaço, ordenando da mais confiável para a mais duvidosa. No caso de medidas planares, use preferencialmente a notação sentido de mergulho/mergulho (Ex: "180/30", "020/40"). Para medidas lineares, utilize mergulho-sentido de mergulho (Ex: "55-340", "70-080"). Use sempre 3 dígitos para o sentido e 2 dígitos para o mergulho.
 
### Como utilizar o software para montar o template da caderneta
- Execute a ferramenta (arquivo .exe ou .elf).
- Clique no botão “Selecionar” e escolha a tabela preenchida nos passos anteriores.
A ferramenta irá analisar se os dados de cada coluna essencial estão no formato correto e mostrará em sua interface. Colunas no formato correto terão o ícone ![ok](https://github.com/FrostPredator/template-builder/assets/114439033/86bfa387-320b-44e7-a71e-f8a474fd1ce2) ao lado enquanto colunas com problemas aparecerão com o ícone ![not_ok](https://github.com/FrostPredator/template-builder/assets/114439033/3e9c5ee1-99d1-4185-b1a9-4e4001d33f09):
 
![interface](https://github.com/FrostPredator/template-builder/assets/114439033/94dbed66-4ac9-4c73-9b26-15895b84f265)

Passar o mouse sobre o ícone revela que tipo de problema está presente na coluna. Também é possível clicar sobre os ícones vermelhos para ver detalhes sobre o problema identificado e em quais linhas, especificamente, ele ocorre:

![popup](https://github.com/FrostPredator/template-builder/assets/114439033/b8b9ef75-3d5f-4834-a239-f26007cbc5e1)

A ferramenta apenas liberará a geração da caderneta quando todos os problemas na tabela forem resolvidos.</br>
Recomenda-se que seja utilizada apenas a tabela fornecida junto à ferramenta para o preenchimento. Utilizar tabelas em outros formatos pode impossibilitar ou limitar as funcionalidades da ferramenta.
- Com todas as colunas devidamente corrigidas, clique no botão “Gerar caderneta” para preencher o template com os dados da tabela. 
- Depois disso, em um editor de texto, basta adicionar as descrições dos afloramentos e amostras, assim como os painéis de croquis e fotos.

<sub>Obs: Devido a diferenças de software, podem haver problemas de formatação caso a caderneta seja editada no Google Docs. Recomenda-se que seja utilizado o Microsoft Office Word ou, no caso de alternativas gratuitas, o ONLYOFFICE ou Softmaker FreeOffice. Para edição colaborativa, a versão online do Word pode ser usada gratuitamente.</sub>

#### Erros comuns durante a execução da ferramenta
##### “Dependência não encontrada: [...]/recursos_app/modelos/template_estilos.docx. Restaure o arquivo a partir do repositório e tente novamente.” (ao abrir a ferramenta)
A ferramenta depende de um arquivo “template_estilos.docx” com estilos pré-definidos para funcionar. Esse arquivo se encontra na pasta recursos_app/modelos, que deve ficar junto ao executável da ferramenta. Caso o arquivo ou a pasta em questão sejam excluídos ou movidos para outro local, ocorrerá esse erro, e basta restaurá-los ao local original para solucioná-lo.

##### Os ícones da interface não estão sendo exibidos.
De forma similar ao erro anterior, basta restaurar os ícones da interface para a pasta recursos_app/icones a partir do arquivo baixado ou do repositório.

##### “ERRO: [Errno 13] Permission denied: [...].docx” (ao salvar a caderneta)
Caso você já tenha gerado a caderneta anteriormente com a ferramenta e esteja gerando um novo arquivo no mesmo caminho, verifique se o arquivo anterior não está aberto em outro programa (Ex: Word). Se não for o caso, tente escolher outra pasta para salvar o arquivo (Ex: Área de trabalho, Downloads, Documentos).
