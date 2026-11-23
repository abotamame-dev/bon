// Lance une session de l'agent Claude "Synthèse hebdo" et récupère ses fichiers de sortie.
import Anthropic, { toFile } from "@anthropic-ai/sdk";
import type { Config } from "./config.js";
import type { FichierExcel } from "./graph.js";

export interface ResultatAgent {
  sessionId: string;
  urlConsole: string;
  verdict: string;
  sorties: Array<{ nom: string; contenu: Buffer }>;
}

// Critères vérifiés par le correcteur automatique. Point de départ : à ajuster
// une fois que vous aurez vu une ou deux présentations.
function rubrique(nomPptx: string, nomControle: string, nbFichiers: number): string {
  return `# Grille de contrôle - synthèse hebdomadaire (version de départ, à ajuster)
- Un seul fichier PowerPoint \`/mnt/session/outputs/${nomPptx}\` existe et s'ouvre sans erreur avec python-pptx.
- Les ${nbFichiers} fichiers Excel de \`/mnt/session/uploads/\` ont tous été lus ; chaque fichier ignoré ou illisible est listé avec la raison dans \`/mnt/session/outputs/${nomControle}\`.
- Le titre de la première diapositive indique la période couverte (date du premier et du dernier fichier).
- Chaque chiffre affiché dans la présentation se retrouve à l'identique dans les données consolidées ; le fichier de contrôle donne pour chacun le calcul utilisé.
- La présentation contient au moins : les totaux de la semaine, l'évolution jour par jour, et les points d'attention (valeurs manquantes, anomalies, fichiers en retard).
- Aucune valeur n'est inventée : les données manquantes sont signalées comme telles.
- Le texte est en français, sans texte provisoire, "TODO" ni zone vide.
`;
}

export async function lancerAgent(config: Config, fichiers: FichierExcel[], libelleSemaine: string): Promise<ResultatAgent> {
  const client = new Anthropic(); // lit ANTHROPIC_API_KEY
  const nomPptx = `synthese-${libelleSemaine}.pptx`;
  const nomControle = `controle-${libelleSemaine}.md`;

  // 1. Envoi des fichiers Excel (montés en lecture seule dans /mnt/session/uploads/).
  const resources = [];
  const noms = new Set<string>();
  const nomsMontes: string[] = [];
  for (const [i, f] of fichiers.entries()) {
    // Préfixe date + numéro : évite les collisions de noms identiques d'un jour à l'autre.
    let nom = `${f.recuLe.slice(0, 10)}_${String(i + 1).padStart(2, "0")}_${f.nom}`;
    while (noms.has(nom)) nom = `x_${nom}`;
    noms.add(nom);
    nomsMontes.push(nom);
    const envoye = await client.files.upload({ file: await toFile(f.contenu, nom) });
    resources.push({ type: "file" as const, file_id: envoye.id, mount_path: `/semaine/${nom}` });
  }

  const inventaire = fichiers
    .map((f, i) => `- ${nomsMontes[i]} (reçu le ${f.recuLe}, de ${f.expediteur}, objet « ${f.objet} »)`)
    .join("\n");

  // 2. Création de la session (l'ID de l'agent et de l'environnement viennent de `ant apply`).
  const session = await client.beta.sessions.create({
    agent: config.agentId,
    environment_id: config.environmentId,
    title: `Synthèse hebdo ${libelleSemaine}`,
    resources,
    budget: { type: "limit", max_list_cost: { amount: config.budgetCentimes, currency: "USD" } },
  });
  const urlConsole = `https://platform.claude.com/workspaces/default/sessions/${session.id}`;
  console.log(`Session ${session.id} — suivi en direct : ${urlConsole}`);

  // 3. Flux d'événements ouvert AVANT l'envoi de la consigne, puis envoi de l'objectif.
  const flux = await client.beta.sessions.events.stream(session.id);
  await client.beta.sessions.events.send(session.id, {
    events: [
      {
        type: "user.define_outcome",
        description:
          `Consolider les ${fichiers.length} fichiers Excel de la semaine (${libelleSemaine}) montés dans ` +
          `/mnt/session/uploads/semaine/ et produire la présentation de synthèse ` +
          `/mnt/session/outputs/${nomPptx} ainsi que le fichier de contrôle /mnt/session/outputs/${nomControle} ` +
          `(fichiers lus, fichiers ignorés et pourquoi, calcul de chaque chiffre, anomalies).\n\n` +
          `Fichiers reçus :\n${inventaire}`,
        rubric: { type: "text", content: rubrique(nomPptx, nomControle, fichiers.length) },
        max_iterations: 4,
      },
    ],
  });

  let verdict = "inconnu";
  boucle: for await (const evt of flux) {
    switch (evt.type) {
      case "agent.message":
        for (const bloc of evt.content) if (bloc.type === "text") process.stdout.write(bloc.text);
        break;
      case "agent.tool_use":
      case "agent.mcp_tool_use":
        // Personne ne surveille l'exécution : toute action que le serveur n'a pas pu juger est refusée.
        if (evt.evaluated_permission === "ask") {
          await client.beta.sessions.events.send(session.id, {
            events: [
              {
                type: "user.tool_confirmation",
                tool_use_id: evt.id,
                result: "deny",
                deny_message: "Exécution automatique sans surveillance : action refusée. Continue sans elle.",
              },
            ],
          });
        }
        break;
      case "span.outcome_evaluation_end":
        verdict = evt.result;
        console.log(`\n[contrôle] itération ${evt.iteration} : ${evt.result} — ${evt.explanation}`);
        break;
      case "session.status_idle":
        if (evt.stop_reason.type !== "requires_action") {
          if (evt.stop_reason.type !== "end_turn") verdict = `arrêt : ${evt.stop_reason.type}`;
          break boucle;
        }
        break;
      case "session.status_terminated":
        verdict = "session terminée";
        break boucle;
    }
  }

  // 4. Récupération des fichiers produits (léger délai d'indexation possible).
  const sorties: ResultatAgent["sorties"] = [];
  for (let essai = 0; essai < 5 && sorties.length === 0; essai++) {
    if (essai > 0) await new Promise((r) => setTimeout(r, 3000));
    for await (const f of client.beta.files.list({ scope_id: session.id, betas: ["managed-agents-2026-04-01"] })) {
      const reponse = await client.beta.files.download(f.id);
      sorties.push({ nom: f.filename, contenu: Buffer.from(await reponse.arrayBuffer()) });
    }
  }

  return { sessionId: session.id, urlConsole, verdict, sorties };
}
