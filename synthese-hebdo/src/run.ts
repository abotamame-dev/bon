// Point d'entrée : récupère les Excel de la semaine, fait produire la présentation par
// l'agent Claude, puis la dépose dans OneDrive et la livre par Outlook.
import { chargerConfig } from "./config.js";
import { Graph, TAILLE_MAX_PIECE_JOINTE } from "./graph.js";
import { lancerAgent } from "./agent.js";

function echapper(texte: string): string {
  return texte.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

async function main() {
  const config = chargerConfig();
  const graph = new Graph(config);

  const maintenant = new Date();
  const depuis = new Date(maintenant.getTime() - config.joursEnArriere * 24 * 3600 * 1000);
  const libelleSemaine = maintenant.toISOString().slice(0, 10);

  console.log(`Recherche des fichiers Excel reçus depuis le ${depuis.toISOString()}…`);
  const fichiers = await graph.recupererFichiersExcel(depuis);
  console.log(`${fichiers.length} fichier(s) trouvé(s).`);

  if (fichiers.length === 0) {
    await graph.livrerParMail({
      objet: `Synthèse hebdo ${libelleSemaine} : aucun fichier reçu`,
      corpsHtml:
        `<p>Aucun fichier Excel correspondant aux filtres n'a été trouvé depuis le ` +
        `${depuis.toLocaleDateString("fr-FR")}. Aucune présentation n'a été générée.</p>`,
    });
    console.log("Aucun fichier : e-mail d'alerte créé, fin.");
    return;
  }

  const resultat = await lancerAgent(config, fichiers, libelleSemaine);
  const pptx = resultat.sorties.find((s) => s.nom.toLowerCase().endsWith(".pptx"));
  const controle = resultat.sorties.find((s) => s.nom.toLowerCase().endsWith(".md"));

  if (!pptx) {
    await graph.livrerParMail({
      objet: `Synthèse hebdo ${libelleSemaine} : échec de génération`,
      corpsHtml:
        `<p>L'agent n'a pas produit de présentation (état : ${echapper(resultat.verdict)}).</p>` +
        `<p>Détail de l'exécution : <a href="${resultat.urlConsole}">${resultat.urlConsole}</a></p>`,
    });
    throw new Error(`Pas de présentation produite (état : ${resultat.verdict}).`);
  }

  let lienOneDrive: string | undefined;
  if (config.dossierOneDrive) {
    lienOneDrive = await graph.deposerOneDrive(config.dossierOneDrive, pptx.nom, pptx.contenu);
    if (controle) await graph.deposerOneDrive(config.dossierOneDrive, controle.nom, controle.contenu);
    console.log(`Déposé dans OneDrive : ${lienOneDrive}`);
  }

  const avertissementControle =
    resultat.verdict === "satisfied"
      ? ""
      : `<p><b>Attention :</b> le contrôle automatique n'a pas validé tous les critères ` +
        `(état : ${echapper(resultat.verdict)}). Relisez la présentation avant diffusion.</p>`;

  await graph.livrerParMail({
    objet: `Synthèse hebdo ${libelleSemaine}`,
    corpsHtml:
      `<p>Bonjour,</p><p>Voici la synthèse de la semaine, construite à partir de ${fichiers.length} fichier(s).</p>` +
      avertissementControle +
      (lienOneDrive ? `<p>Présentation dans OneDrive : <a href="${lienOneDrive}">${echapper(pptx.nom)}</a></p>` : "") +
      (pptx.contenu.length > TAILLE_MAX_PIECE_JOINTE && !lienOneDrive
        ? `<p>La présentation est trop lourde pour être jointe ; activez DOSSIER_ONEDRIVE pour recevoir un lien.</p>`
        : "") +
      (controle ? `<pre>${echapper(controle.contenu.toString("utf8"))}</pre>` : ""),
    pieceJointe: pptx,
  });

  console.log(`Livraison faite (${config.modeLivraison}). Verdict du contrôle : ${resultat.verdict}.`);
}

main().catch((erreur) => {
  console.error(erreur);
  process.exit(1);
});
