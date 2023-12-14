function addCadastro() {
  // Obter os valores dos inputs
  var matriculaValue = document.getElementById("matricula").value;
  var cpfValue = document.getElementById("cpf").value;
  var identidadeValue = document.getElementById("identidade").value;
  var orgaoEmissorValue = document.getElementById("orgao_emissor").value;
  var cnpjValue = document.getElementById("cnpj").value;
  var nomeValue = document.getElementById("nome").value;
  var nacionalidadeValue = document.getElementById("nacionalidade").value;
  var estadoCivilValue = document.getElementById("estado_civil").value;
  var regimeBensSelect = document.getElementById("regimedebens");
  var regimeBensValue =
    regimeBensSelect.options[regimeBensSelect.selectedIndex].text;
  var profissaoValue = document.getElementById("profissao").value;
  var qualificacaoValue = document.getElementById("qualificacao").value;
  var ufValue = document.getElementById("uf").value;
  var cepValue = document.getElementById("cep").value;
  var tipoLogradouroValue = document.getElementById("tipo_logradouro").value;
  var logradouroValue = document.getElementById("logradouro").value;
  var numeroValue = document.getElementById("numero").value;
  var unidadeValue = document.getElementById("unidade").value;
  var cidadeValue = document.getElementById("cidade").value;
  var bairroValue = document.getElementById("bairro").value;
  var complementoValue = document.getElementById("complemento").value;
  var nomePaiValue = document.getElementById("nome_pai").value;
  var nomeMaeValue = document.getElementById("nome_mae").value;

  // Criar uma nova linha na tabela
  var table = document
    .getElementById("cadastrosTable")
    .getElementsByTagName("tbody")[0];
  var newRow = table.insertRow(table.rows.length);

  // Preencher a nova linha com os valores dos inputs
  newRow.insertCell().innerHTML = matriculaValue;
  newRow.insertCell().innerHTML = cpfValue.toUpperCase();
  newRow.insertCell().innerHTML = identidadeValue.toUpperCase();
  newRow.insertCell().innerHTML = orgaoEmissorValue.toUpperCase();
  newRow.insertCell().innerHTML = cnpjValue.toUpperCase();
  newRow.insertCell().innerHTML = nomeValue.toUpperCase();
  newRow.insertCell().innerHTML = nacionalidadeValue.toUpperCase();
  newRow.insertCell().innerHTML = estadoCivilValue.toUpperCase();
  newRow.insertCell().innerHTML = regimeBensValue.toUpperCase();
  newRow.insertCell().innerHTML = profissaoValue.toUpperCase();
  newRow.insertCell().innerHTML = qualificacaoValue.toUpperCase();
  newRow.insertCell().innerHTML = ufValue.toUpperCase();
  newRow.insertCell().innerHTML = cepValue.toUpperCase();
  newRow.insertCell().innerHTML = tipoLogradouroValue.toUpperCase();
  newRow.insertCell().innerHTML = logradouroValue.toUpperCase();
  newRow.insertCell().innerHTML = numeroValue.toUpperCase();
  newRow.insertCell().innerHTML = unidadeValue.toUpperCase();
  newRow.insertCell().innerHTML = cidadeValue.toUpperCase();
  newRow.insertCell().innerHTML = bairroValue.toUpperCase();
  newRow.insertCell().innerHTML = complementoValue.toUpperCase();
  newRow.insertCell().innerHTML = nomePaiValue.toUpperCase();
  newRow.insertCell().innerHTML = nomeMaeValue.toUpperCase();
}

function saveToExcel() {
  // Criar uma matriz com os dados da tabela
  var data = [
    [
      "N° MATRICULA",
      "CPF",
      "IDENTIDADE",
      "ÓRGÃO EM",
      "CNPJ",
      "NOME",
      "NACIONALIDADE",
      "ESTADO CÍVIL",
      "REGIME DE BENS",
      "PROFISSÃO",
      "QUALIFICAÇÃO",
      "UF",
      "CEP",
      "TIPO DE LOGRADOURO",
      "LOGRADOURO",
      "NÚMERO",
      "UNIDADE",
      "CIDADE",
      "BAIRRO",
      "COMPLEMENTO",
      "NOME DE PAI",
      "NOME DA MÃE",
    ],
  ]; // Cabeçalho

  var table = document
    .getElementById("cadastrosTable")
    .getElementsByTagName("tbody")[0];
  var rows = table.getElementsByTagName("tr");
  for (var i = 0; i < rows.length; i++) {
    var cells = rows[i].getElementsByTagName("td");
    var rowData = [];
    for (var j = 0; j < cells.length; j++) {
      rowData.push(cells[j].innerText);
    }
    data.push(rowData);
  }

  // Criar um objeto de workbook
  var wb = XLSX.utils.book_new();

  // Adicionar uma planilha ao workbook
  var ws = XLSX.utils.aoa_to_sheet(data);
  XLSX.utils.book_append_sheet(wb, ws, "Cadastros");

  // Salvar o arquivo Excel
  XLSX.writeFile(wb, "numero_da_matricula.xlsx");
}
function validarFormulario() {
  var form = document.getElementById("myForm");
  var elementosDoFormulario = form.elements;

  for (var i = 0; i < elementosDoFormulario.length; i++) {
    if (elementosDoFormulario[i].type !== "button") {
      if (elementosDoFormulario[i].value === "") {
        alert("Por favor, preencha todos os campos do formulário.");
        return false;
      }
    }
  }

  // Se todos os campos estão preenchidos, retorna true
  return true;
}
