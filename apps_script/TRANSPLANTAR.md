# Como transplantar tudo para outra conta Google

O site (GitHub Pages) não guarda nada seu. Quem guarda é um **script do Google (Apps Script) na sua conta**: ele cria a pasta "Ferramenta PDPAS" no seu Drive, o arquivo de backup e a "Planilha mestra". Nada disso tem nome de conta, código ou ID escrito no programa. Para mudar de conta, repita os passos abaixo na conta nova.

## Na conta nova (cerca de 10 minutos)
1. Entre em https://script.google.com com a conta nova e escolha **Novo projeto**.
2. Apague o código que vem e cole o conteúdo de `Code.gs` (neste mesmo repositório, pasta `apps_script`). Salve.
3. No menu **Executar**, escolha a função **preparar** e clique em Executar. O Google pede autorização uma vez: é a conta nova dando permissão ao próprio script. Isso cria a pasta e a planilha.
4. Em **Configurações do projeto > Propriedades do script**, adicione a propriedade **CODIGO** com uma senha longa (12 letras ou mais) que só você saiba.
5. **Implantar > Nova implantação > Tipo: App da Web.** Executar como: **Eu**. Quem pode acessar: **Qualquer pessoa**. Implante e copie o endereço que termina em `/exec`.

## No site
1. Abra o site, toque em **Levar para outro computador**.
2. Cole o endereço novo e o código novo. Toque em **Conectar e testar**.
3. Antes disso, nos aparelhos onde estão os pedidos, use **Enviar tudo agora**: ele leva os pedidos e a planilha para a conta nova.
4. Marque **Salvar sozinho a cada alteração**.

## Levar o que já existe na conta antiga
- No Drive antigo, a pasta "Ferramenta PDPAS" pode ser movida ou ter a propriedade transferida (Compartilhar > transferir propriedade), se as duas contas permitirem.
- Ou: no site conectado à conta antiga toque em **Trazer do Drive**, troque o endereço e o código para os da conta nova e toque em **Enviar tudo agora**.

## Segurança
- Quem souber o endereço e o código acessa só o que o script mexe: a pasta e a planilha. Não use um código que já use em outro lugar.
- Para trocar o código: altere a propriedade CODIGO no script e digite o novo no site.
- Para cortar o acesso de vez: Implantar > Gerenciar implantações > arquivar.
