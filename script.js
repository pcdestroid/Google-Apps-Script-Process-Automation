async function executarFuncoes() {

  try {
    let f;
    Logger.log('Verificando arquivo no email para salvar na minha pasta...'); f = await getReport();

    if (f == true) {
      Logger.log('Movendo arquivo da minha pasta para pasta CONSULTA...'); f = await fileMove('SCSC301.XLSX');
    }

    if (f == true) {
      Logger.log('Pegar arquivo XLSX da pasta de CONSULTA, converter para GS e salvar na minha pasta...'); f = await xlsxToGs();
    }

    if (f == true) {
      Logger.log('Movendo arquivo GS da minha pasta para pasta CONSULTA...'); f = await fileMove('REL301');
    }

    if (f == true) {
      Logger.log('Analisando as nota fiscal recebidas NFxSC...'); f = await nfxsc();
    }

    if (f == true) {
      Logger.log('Atualizar a situação das OCs e SCs com base no relatório 301...'); f = await atualizarSit();
    }

    if (f == true) {
      Logger.log('Enviando mensagem para chat "Relatório foi atualizado"... '); f = await sendMessage(`O relatório foi atualizado com sucesso.`);
    }

    if (f == true) {
      console.log("As funções foram executadas com sucesso!");
    }

  } catch (e) {
    console.log("Ocorreu um erro durante a execução das funções:", e);
    sendMessage('Ocorreu um erro no processamento do relatório.')
  }
}

//Pegar o relatório 301 no email e salva na minha pasta
async function getReport() {

  const assuntoAProcurar = 'Arquivo(s) anexo(s).';
  const minhaPastaId = "1O4ZFKMlPy6kI_zqJgxT6hT5OSz5wEq6Q";
  const minhaPasta = DriveApp.getFolderById(minhaPastaId);

  // procurar apenas uma thread com o assunto específico
  const threads = GmailApp.search(`subject:"${assuntoAProcurar}"`, 0, 1);


  if (threads.length > 0 && threads[0].isUnread()) { // se houver pelo menos uma thread encontrada e não foi lido.

    sendMessage(`${bomDia()}! ${msgReceiveReport()}`)
    threads[0].markRead()

    const messages = threads[0].getMessages();
    const attachments = messages[messages.length - 1].getAttachments(); // obter anexos da última mensagem da thread
    for (let j = 0; j < attachments.length; j++) {
      const anexo = attachments[j];
      const nomeDoArquivo = anexo.getName();
      // remover o arquivo existente, se houver
      const arquivosAntigos = minhaPasta.getFilesByName('SCSC301.XLSX');
      while (arquivosAntigos.hasNext()) {
        const arquivoAntigo = arquivosAntigos.next();
        arquivoAntigo.setTrashed(true);
      }
      // salvar o novo anexo na pasta especificada
      minhaPasta.createFile(anexo).setName('SCSC301.XLSX');
      return true;
    }
  } else {
    Logger.log('Nenhum email encontrado...')
    return false;
  }
}

//Mover arquivo de uma pasta para outra pasta
async function fileMove(a) {
  const nomeArquivo = a;
  const minhaPastaId = "1O4ZFKMlPy6kI_zqJgxT6hT5OSz5wEq6Q"; //Minha pasta
  const pastaDestinoId = "1upBCebUcIPZVUA8Ch5YfPS7YXZQyJ6zc"; //Pasta compartilhada
  const minhaPasta = DriveApp.getFolderById(minhaPastaId);
  const pastaDestino = DriveApp.getFolderById(pastaDestinoId);
  const arquivos = minhaPasta.getFilesByName(nomeArquivo);

  if (arquivos.hasNext()) {
    const arquivo = arquivos.next();
    const arquivosDestino = pastaDestino.getFilesByName(nomeArquivo);

    if (arquivosDestino.hasNext()) {
      const arquivoDestino = arquivosDestino.next();
      arquivoDestino.setTrashed(true);
    }

    pastaDestino.addFile(arquivo);
    minhaPasta.removeFile(arquivo);
    return true

  } else {

    Logger.log(`Não foi encontrado nenhum arquivo com o nome ${nomeArquivo} na pasta ${minhaPasta.getName()}.`);
    return false
  }
}

//Converter arquivo XLSX(Planilha excel) em GS(Planilha Google)
async function xlsxToGs() {
  const pastaDestinoId = "1upBCebUcIPZVUA8Ch5YfPS7YXZQyJ6zc"; //Pasta compartilhada
  const pastaDestino = DriveApp.getFolderById(pastaDestinoId);
  const minhaPastaId = "1O4ZFKMlPy6kI_zqJgxT6hT5OSz5wEq6Q"; //Minha pasta
  const minhaPasta = DriveApp.getFolderById(minhaPastaId);

  // Verificar se existe o arquivo SCSC301.XLSX na pasta compartilhada
  const arquivosDestino = pastaDestino.getFilesByName('SCSC301.XLSX');
  if (arquivosDestino.hasNext()) {
    const arquivoDestino = arquivosDestino.next();
    const blob = arquivoDestino.getBlob();
    // Converter relatório 301 em arquivo Google Sheets e salvando na Minha pasta
    const novoArquivo = {
      title: 'REL301',
      parents: [{ id: minhaPastaId }]
    };
    Drive.Files.insert(novoArquivo, blob, { convert: true });
    return true;
  } else {
    return false;
  }
}

// Retorna a frase, bom dia, boa tarde ou boa noite
function bomDia() {
  let frase = '';
  let data = new Date();
  let hora = data.getHours();
  if (hora >= 3 && hora < 12)
    frase = 'Bom dia';
  else if (hora >= 12 && hora < 18)
    frase = 'Boa tarde';
  else if (hora >= 18 && hora <= 23)
    frase = 'Boa noite';
  return frase;
}

function escolherAleatorio(lista) {
  const indiceAleatorio = Math.floor(Math.random() * lista.length);
  return lista[indiceAleatorio];
}

function getMarinas() {
  //Retorna uma lista com nome das Marinas
  let pedidosDados = SpreadsheetApp.openById('1C0DBuuXcuTDPGo2hYZJUfxEHh-mcR_hA5uf7rgATobs').getSheetByName('Dados');
  let marinas = new Set(pedidosDados.getRange(3, 3, pedidosDados.getLastRow() - 2, 1).getValues().map(marina => marina[0]));
  return Array.from(marinas);
}

function atualizarSit() {
  //Atualizar a situação das OCs e SCs com base no relatório 301.
  let marinas = getMarinas();
  let s = SpreadsheetApp.openById('1C0DBuuXcuTDPGo2hYZJUfxEHh-mcR_hA5uf7rgATobs');
  let ul = Number(s.getLastRow());
  Logger.log('Pegar relatório')
  let rel301 = SpreadsheetApp.openById(getIdRel301()).getSheetByName('Plan1');
  let rel301Mar = rel301.getRange(1, 2, rel301.getLastRow(), 1).getValues();
  let rel301Ocs = rel301.getRange(1, 27, rel301.getLastRow(), 1).getValues();
  let rel301Scs = rel301.getRange(1, 3, rel301.getLastRow(), 1).getValues();
  let rel301Sit = rel301.getRange(1, 40, rel301.getLastRow(), 1).getValues();
  let rel301Ssc = rel301.getRange(1, 23, rel301.getLastRow(), 1).getValues();

  for (let t = 0; t < marinas.length; t++) {
    let sheet = s.getSheetByName(marinas[t]);
    let marina = marinas[t];
    let scs = sheet.getRange(4, 2, ul - 3, 1).getValues().toString().split(',');
    let com = sheet.getRange(4, 8, ul - 3, 1).getValues().toString().split(',');
    let ocs = sheet.getRange(4, 9, ul - 3, 1).getValues().toString().split(',');
    let map = sheet.getRange(4, 10, ul - 3, 1).getValues().toString().split(',');
    for (let i = 0; i < scs.length; i++) {
      if (ocs[i] == '') {
        for (let x = 0; x < rel301Scs.length; x++) {
          if (marina + scs[i] == rel301Mar[x] + rel301Scs[x]) {
            sheet.getRange(i + 4, 12).setValue(`SC ${rel301Ssc[x].toString().replace('ado', 'ada')}`);
            break;
          }
        }
      } else {
        for (let x = 0; x < rel301Ocs.length; x++) {
          if (marina + ocs[i] == rel301Mar[x] + rel301Ocs[x]) {
            if (rel301Sit[x] == "Aprovado" && sheet.getRange(i + 4, 12).getValue() !== "Aprovada" && sheet.getRange(i + 4, 3).getValue() !== "COMPRADO" && sheet.getRange(i + 4, 3).getValue() !== "ENTREGUE") {

              sendMessage(`${com[i]}, a OC${ocs[i]} na ${marinas[t]} foi aprovada!`)
            }
            sheet.getRange(i + 4, 12).setValue(rel301Sit[x].toString().replace('ado', 'ada'));
            break;
          }
        }
      }
    }
  }
  return true
}

//Envia mensagem no Google Chat
function sendMessage(text) {
  const payload = JSON.stringify({ text: text });
  const options = {
    method: 'POST',
    contentType: 'application/json',
    payload: payload,
  };
  UrlFetchApp.fetch(GOOGLE_CHAT_WEBHOOK_LINK, options);
  return true
}

function getIdRel301() {
  //Retorna id do relatório 301 que consta na pasta consulta.
  let pastaConsulta = DriveApp.getFolderById("1upBCebUcIPZVUA8Ch5YfPS7YXZQyJ6zc");
  let arquivos = pastaConsulta.searchFiles('mimeType="application/vnd.google-apps.spreadsheet" and title="REL301"'); // Pesquise por arquivos do tipo planilha com o título "REL301"
  while (arquivos.hasNext()) {
    let nomeArquivo = arquivos.next();
    let id = nomeArquivo.getId();
    return id;
  }
}

function msgReceiveReport() {
  let dados = SpreadsheetApp.openById('1ziJX3KK9D-0aRDTqJ9-xieeUO8fGkGcEWHwxm3Z18Bc').getSheetByName('Dados')
  let msgs = dados.getRange(2, 6, 5, 1).getValues()
  let lista = [];
  for (let i = 0; i < msgs.length; i++) {
    lista.push(msgs[i][0]);
  }
  return escolherAleatorio(lista)
}

function nfxsc() {
  const now = new Date();
  let thisworkbook = SpreadsheetApp.openById('1ziJX3KK9D-0aRDTqJ9-xieeUO8fGkGcEWHwxm3Z18Bc');
  let registroNFxSC = thisworkbook.getSheetByName("NFxSC");
  let dados = thisworkbook.getSheetByName("Dados");
  let abrev = dados.getRange(2, 1, dados.getLastRow() - 1, 1).getValues().toString().split(',');
  let nome = dados.getRange(2, 2, dados.getLastRow() - 1, 1).getValues().toString().split(',');
  let qtdComp = Number(registroNFxSC.getRange(1, 6).getValue());
  let regLastRow = registroNFxSC.getLastRow() + 1

  let tabela_arquivo_processados = registroNFxSC.getRange(2, 2, registroNFxSC.getLastRow() - 1, 1).getValues()

  let pastabackup = DriveApp.getFolderById("1spbRB4Rgo8gDJUKTnboMM8jUkYAVV1ck"); //Pasta para salva backup das notas fiscais
  let idRel = getIdRel301()

  let listahtml = '';
  let qtdNfs = 0, qtdNfsAp = 0;

  //Pastas dos compradores.
  let pastaComprador = [
    { comprador: 'comprador1', id: '1dEpQimv6i_EfnZVInNJzvaW8CLdAz5KO' },
    { comprador: 'comprador2', id: '1pYNyqmOJwL7L2GCwbrPeFMmqgDO2_In8' },
    { comprador: 'comprador3', id: '1L1sqPuuLzoO3Db2laZMOddVr6M3cdfms' },
    { comprador: 'comprador4', id: '10cYPWk9Np5OmU0XL2DWllBARrp8g-6cD' },
    { comprador: 'comprador5', id: '1sZ8h63cWG0ncp9C_XUdcY2gsAZo3w9-f' }
  ];

  //Pegar o relatório 301 na pasta
  let rel301 = SpreadsheetApp.openById(idRel).getSheetByName('Plan1');
  let rel301Data = DriveApp.getFileById(idRel).getDateCreated();

  let allData = {}
  let values = rel301.getRange(2, 1, rel301.getLastRow() - 1, 27).getValues();
  values.map(row => {

    let empresa = row[1]
    let pedido = row[2].toString()
    let situacao = row[22]

    if (empresa !== "Empresa") {
      if (!allData[empresa]) {
        allData[empresa] = { pedidos: [] }
      }
      let pedidoIndex = allData[empresa].pedidos.findIndex(p => p.numero === pedido)
      if (pedidoIndex === -1) {
        allData[empresa].pedidos.push({ numero: pedido, situacoes: [situacao] })
      } else {
        allData[empresa].pedidos[pedidoIndex].situacoes.push(situacao)
      }
    }
  }).join('');

  let pastaNotas = DriveApp.getFolderById("0ABX_Ev8dn-ofUk9PVA"); // Pasta com as nota fiscais com solicitação de compra
  let notas = pastaNotas.getFiles(); // Todas notas da pasta

  while (notas.hasNext()) {
    let arquivo = notas.next();
    let nomeArquivo = arquivo.getName();
    let nomeMarina;

    //Validações da renomeação do arquivo
    let verNF, verSC, verSPC, verAbrev;

    if (nomeArquivo.indexOf(' ') >= 0) { // Se existe espaço no arquivo: Remove todos os espaços e renomeia o arquivo
      arquivo.setName(nomeArquivo.replace(/ /g, ""));
      nomeArquivo = arquivo.getName();
      //Logger.log(`${nomeArquivo} foi renomeado devido a espaços desnecessários`);
      listahtml = listahtml + `<br><span style="color:red;"><strong>Aviso: ${nomeArquivo} foi renomeado devido a espaços desnecessários</strong></span>`;
    }
    verSPC = nomeArquivo.indexOf(' ') >= 0

    verNF = nomeArquivo.indexOf('-NF') >= 0
    if (verNF) { //Verifica -NF
      //Logger.log(`${nomeArquivo} nome do arquivo tem "-NF"`);
    } else {
      //Logger.log(`${nomeArquivo} Erro: "-NF"`);
      listahtml = listahtml + `<br><span style="color:red;"><strong>Erro: "-NF" ${nomeArquivo}</strong></span>`;
    }

    verSC = nomeArquivo.indexOf('-SC') >= 0
    if (verSC) { // Verifica -SC
      //Logger.log(`${nomeArquivo} nome do arquivo tem "-SC"`);
    } else {
      //Logger.log(`${nomeArquivo} Erro: "-SC"`);
      listahtml = listahtml + `<br><span style="color:red;"><strong>Erro: "-SC" ${nomeArquivo}</strong></span>`;
    }

    let arq = nomeArquivo.toString().replace('NF', '').replace('SC', '').replace('.pdf', '').replace('.PDF', '').split('-')

    for (let i = 0; i < abrev.length; i++) {
      if (abrev[i] == arq[0]) {
        nomeMarina = nome[i]
        verAbrev = true; break;
      } else {
        verAbrev = false
      }
    }

    if (verAbrev == false) {
      //Logger.log(`${nomeArquivo} abreviação errada`);
      listahtml = listahtml + `<br><span style="color:red;"><strong>Erro: abreviação da marina ${nomeArquivo}</strong></span>`
    }

    if (!verSPC && verNF && verSC && verAbrev) {

      //Validação da situação de cada SC
      let verSit;

      Logger.log(arq)

      let x = allData[nomeMarina].pedidos.find(p => p.numero === arq[2])

      Logger.log(x)
      Logger.log(nomeMarina)
      Logger.log(!verSPC)
      Logger.log(verNF)
      Logger.log(verSC)
      Logger.log(verAbrev)

      if (x != null) {

        for (let t = 0; t < x.situacoes.length; t++) {
          if (x.situacoes[t] == "Aprovado") {
            Logger.log(`${nomeArquivo} Aprovado`)
            verSit = true
          }

          if (x.situacoes[t] == "Reprovado") {
            Logger.log(`${nomeArquivo} Reprovado`)
            listahtml = listahtml + `<br><span style="color:red;"><strong>Erro: SC foi reprovada ${nomeArquivo}</strong></span>`
            verSit = false
            break
          }
          if (x.situacoes[t] == "Em análise") {
            Logger.log(`${nomeArquivo} Em análise`)
            listahtml = listahtml + `<br><span>Não aprovada: ${nomeArquivo}</span>`
            verSit = false
            break
          }

        }
        if (verSit) {

          //verificar que o arquivo em algum momento já foi processado anteriomente
          for (let i = 0; i < tabela_arquivo_processados.length; i++) {
            if (tabela_arquivo_processados[i][0] == nomeArquivo) {
              listahtml = listahtml + `<br><span style="color:red;"><strong>Aviso: ${nomeArquivo} já foi processado anteriomente.</strong></span>`;
              break
            }
          }

          Logger.log(`${nomeArquivo} Tudo foi aprovado`)
          if (qtdComp < 0) { qtdComp = 4; }
          let pastaDestino = DriveApp.getFolderById(pastaComprador[qtdComp].id);
          Logger.log(`Movendo arquivo ${nomeArquivo} para pasta ${pastaComprador[qtdComp].comprador}`);

          arquivo.makeCopy(arquivo.getName(), pastabackup);

          let copy = pastabackup.getFilesByName(arquivo.getName()).next();
          if (copy) {
            arquivo.moveTo(pastaDestino);
            registroNFxSC.getRange(regLastRow, 1).setValue(now);
            registroNFxSC.getRange(regLastRow, 2).setValue(nomeArquivo);
            registroNFxSC.getRange(regLastRow, 3).setValue(`Arquivo movido para pasta do(a) ${pastaComprador[qtdComp].comprador}`);
            listahtml = listahtml + `<br><span style="color:green;"><strong>Aprovada: ${nomeArquivo} - movido para pasta - ${pastaComprador[qtdComp].comprador}</strong></span>`
            regLastRow++
          }
          qtdNfsAp++
          qtdComp = qtdComp - 1

        }
      } else {
        listahtml = listahtml + `<br><span style="color:red;"><strong>Erro: SC não encontrada no relatório ${nomeArquivo}</strong></span>`
      }

    } else {
      Logger.log(`Erro: Verifique o arquivo ${nomeArquivo}`)
    }
    qtdNfs++
  }

  Logger.log('Total de nota fiscais: ' + qtdNfs)
  Logger.log('Notas Fiscais com SC aprovada: ' + qtdNfsAp)
  Logger.log('Notas fiscais com SC não aprovada: ' + (qtdNfs - qtdNfsAp))

  let subtotal = `<span><strong>Total de notas fiscais: ${qtdNfs}</strong></span><br><span><strong>Notas Fiscais com SC aprovada: ${qtdNfsAp}</strong></span><br><span><strong>Notas fiscais com SC não aprovada: ${(qtdNfs - qtdNfsAp)}</strong></span>`

  registroNFxSC.getRange(1, 6).setValue(Number(qtdComp))

  let email = 'compras@yourdomain.com.br';
  let assunto = 'Análise - NFxSC'
  emailTemp = HtmlService.createHtmlOutput(`<span>${bomDia()}!</span><br><span>Segue abaixo a relação dos arquivo de notas fiscais com solicitações de compra.</span><br>${listahtml}<br><br>${subtotal}<br><br><span>Att.,</span><br><span><strong>Compras - ###</strong></span><br><span><img src="https://intranet.####.com.br/wp-content/uploads/2020/05/logo.png" width="136" height="48"></span>`).setTitle('Email');

  let htmlMessage = emailTemp.getContent()
  GmailApp.sendEmail(email, assunto, "", { 'from': 'compras@yourdomain.com.br', name: 'NFxSC', htmlBody: htmlMessage });

  if (qtdNfsAp == 1) { sendMessage(`${qtdNfsAp} nota foi aprovada.`) } else {
    if (qtdNfsAp > 1) { sendMessage(`${qtdNfsAp} notas foram aprovadas e já foram movidas para pasta de vocês.`) }
  }

  return true

}
