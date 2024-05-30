/**
 * Função para procurar emails com um determinado assunto e registrar as informações relevantes no console.
 */
function procurarEmail() {
  // ID da planilha
  var planilhaId = "1KwQn7avnsB6zbsgp1hKdvLG8WkTSe3OYpP5Tv5QiQw8";
  
  // Abre a planilha pelo ID
  var planilha = SpreadsheetApp.openById(planilhaId);
  
  // Obtém a célula A5 na planilha "teste1"
  var valorCelula = planilha.getSheetByName("teste1").getRange("A5");
  
  // Obtém o valor da célula A5 e remove espaços em branco extras
  var assunto = valorCelula.getValue().trim();
  
  // Registra o valor da célula A5 no console
  Logger.log("Célula A5: " + assunto);
  
  // Faz a autenticação com a API do Gmail
  var accessToken = ScriptApp.getOAuthToken();
  
  // Pesquisa no Gmail com o assunto especificado
  var threads = GmailApp.search('subject:"' + assunto + '"');
  
  // Verifica se há threads encontradas
  if (threads.length > 0) {
    Logger.log("Emails encontrados: " + threads.length);
    
    // Processa cada thread encontrada
    for (var i = 0; i < threads.length; i++) {
      var thread = threads[i];
      var messageCount = thread.getMessageCount(); // Obtém o número de mensagens na thread
      Logger.log("Mensagens encontradas: " + messageCount);

      // Itera sobre cada mensagem na thread
      for (var j = 0; j < messageCount; j++) {
        var msgNumero = j + 1;
        var message = thread.getMessages()[j]; // Obtém a mensagem na posição j        
        // Extrai as informações da mensagem
        var subject = message.getSubject(); // Obtém o assunto da mensagem
        var from = message.getFrom(); // Obtém o remetente da mensagem
        var date = message.getDate(); // Obtém a data da mensagem
        var body = message.getPlainBody(); // Obtém o corpo da mensagem em texto simples          
        
        // Remove os caracteres >>>
        body = body.replace(/>/g, '');
        
        // Registra as informações relevantes no console
        Logger.log("#####################");
        Logger.log("Mensagem nº: " + msgNumero);
        Logger.log("Remetente: " + from);
        Logger.log("Assunto: " + subject);
        Logger.log("Data: " + date);
        Logger.log("Corpo: " + body);
        
        // Adicione o código para escrever essas informações na planilha, se necessário
      }
    }
  } else {
    Logger.log("Nenhuma thread encontrada com o assunto: " + assunto);
  }
}
