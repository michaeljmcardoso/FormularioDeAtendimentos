document.addEventListener('DOMContentLoaded', () => {
    const form = document.getElementById('meu_form');
    const btnEnviar = document.getElementById('bt-enviar');
    const btnExportar = document.getElementById('bt-exportar');

    btnEnviar.addEventListener('click', (event) => {
        event.preventDefault();
        salvarForm();
    });

    btnExportar.addEventListener('click', (event) => {
        event.preventDefault();
        exportToExcel();
    });

    function salvarForm() {
        const nome = document.getElementById('nome').value;
        const interessado = document.getElementById('interessado').value;
        const comunidade = document.getElementById('comunidade').value;
        const fone = document.getElementById('fone').value;
        const email = document.getElementById('email').value;
        const protocolo = document.getElementById('protocolo').value;
        const data = document.getElementById('data').value;
        const motivoatendimento = document.getElementById('motivoatendimento').value;

        if (!nome) {
            alert('Por favor, preencha o campo nome.');
            return;
        }

        if (localStorage.cont) {
            localStorage.cont = Number(localStorage.cont) + 1;
        } else {
            localStorage.cont = 1;
        }

        const cad = `${nome};${interessado};${comunidade};${fone};${email};${protocolo};${motivoatendimento};${data}`;
        localStorage.setItem(`cad_${localStorage.cont}`, cad);

        alert(`Obrigado sr(a) ${nome}, os seus dados foram encaminhados com sucesso!`);

        form.reset();
    }

    function exportToExcel() {
        const data = [];
        for (let i = 1; i <= localStorage.cont; i++) {
            const item = localStorage.getItem(`cad_${i}`);
            if (item) {
                data.push(item.split(';'));
            }
        }

        const planilha = XLSX.utils.aoa_to_sheet([
            ["Nome", "Interessado", "Comunidade", "Fone", "Email", "Protocolo", "Motivo do Atendimento", "Data"],
            ...data
        ]);

        const pastaDeTrabalho = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(pastaDeTrabalho, planilha, "Atendimentos");

        XLSX.writeFile(pastaDeTrabalho, 'atendimentos.xlsx');
    }
});
