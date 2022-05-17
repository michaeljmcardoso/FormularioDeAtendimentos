function Enviar() {

    var nome = document.getElementById("nome");

    if (nome.value != "") {
        alert('Obrigado sr(a) ' + nome.value + ' os seus dados foram encaminhados com sucesso!');
    }

}