const fs = require("fs");
const XLSX = require("xlsx");

const pasta = "./backups";

const arquivos = fs.readdirSync(pasta)
    .filter(f => f.endsWith(".json"))
    .sort();

if (!arquivos.length) {
    throw new Error("Nenhum backup encontrado.");
}

const ultimo = arquivos[arquivos.length - 1];

const caminho = `${pasta}/${ultimo}`;

const backup = JSON.parse(fs.readFileSync(caminho, "utf8"));

const wb = XLSX.utils.book_new();

const mapa = {};

backup.streamers.forEach(s => {
    mapa[s.id] = s.nome;
});

//////////////////////////////////////////////////////
// STREAMERS
//////////////////////////////////////////////////////

const streamers = backup.streamers.map(s => ({
    ID: s.id,
    Nome: s.nome,
    Twitch: s.twitch_id || "",
    Kick: s.kick_id || "",
    Status: s.status,
    Parceiro: s.is_partner ? "SIM" : "NÃO",
    Criado: s.created_at,
    Atualizado: s.updated_at
}));

XLSX.utils.book_append_sheet(
    wb,
    XLSX.utils.json_to_sheet(streamers),
    "Streamers"
);

//////////////////////////////////////////////////////
// SESSÕES
//////////////////////////////////////////////////////

const sessoes = backup.streamer_sessions.map(s => {

    let duracao = "";

    if (s.saida) {

        const inicio = new Date(s.entrada);
        const fim = new Date(s.saida);

        const minutos = Math.floor((fim - inicio) / 60000);

        duracao =
            `${Math.floor(minutos / 60)}h ${minutos % 60}m`;
    }

    return {

        ID: s.id,

        Streamer: mapa[s.streamer_id] || s.streamer_id,

        Entrada: s.entrada,

        Saída: s.saida,

        Duração: duracao

    };

});

XLSX.utils.book_append_sheet(
    wb,
    XLSX.utils.json_to_sheet(sessoes),
    "Sessões"
);

//////////////////////////////////////////////////////
// ESTATÍSTICAS
//////////////////////////////////////////////////////

const stats = [

    {
        Item: "Total de Streamers",
        Valor: backup.streamers.length
    },

    {
        Item: "Parceiros",
        Valor: backup.streamers.filter(s => s.is_partner).length
    },

    {
        Item: "Online",
        Valor: backup.streamers.filter(s => s.status === "online").length
    },

    {
        Item: "Offline",
        Valor: backup.streamers.filter(s => s.status === "offline").length
    },

    {
        Item: "Sessões",
        Valor: backup.streamer_sessions.length
    }

];

XLSX.utils.book_append_sheet(
    wb,
    XLSX.utils.json_to_sheet(stats),
    "Estatísticas"
);

const nomeExcel = caminho.replace(".json", ".xlsx");

XLSX.writeFile(wb, nomeExcel);

console.log(nomeExcel);
