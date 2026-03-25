async function iniciarConsulta() {
    const payload = {
        numeros: document.getElementById("numeros").value,
        usuario: document.getElementById("usuario").value,
        senha: document.getElementById("senha").value
    };

    const resp = await fetch("/consultar", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload)
    });

    const data = await resp.json();
    acompanharStatus(data.job_id);
    document.getElementById("statusBox").style.display = "block";
}

async function acompanharStatus(job_id) {
    const intervalo = setInterval(async () => {
        const r = await fetch(`/status/${job_id}`);
        const d = await r.json();

        document.getElementById("msgStatus").innerText =
            `${d.pct}% - ${d.msg}`;

        if (d.status === "done") {
            clearInterval(intervalo);
            window.location.href = `/download/${job_id}`;
        }
        if (d.status === "error") {
            clearInterval(intervalo);
            alert("Erro: " + d.msg);
        }
    }, 2000);
}