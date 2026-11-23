// Configuration lue depuis les variables d'environnement.
// En local : copiez `.env.example` vers `.env` et complétez les valeurs `A_COMPLETER`.
// Sur GitHub : secrets et variables du dépôt (voir README.md).

function lire(nom: string, defaut?: string): string {
  const valeur = process.env[nom]?.trim() || defaut;
  if (valeur === undefined || valeur === "" || valeur.startsWith("A_COMPLETER")) {
    throw new Error(`Variable d'environnement manquante ou à compléter : ${nom}`);
  }
  return valeur;
}

function lireOptionnel(nom: string): string | undefined {
  const valeur = process.env[nom]?.trim();
  return valeur && !valeur.startsWith("A_COMPLETER") ? valeur : undefined;
}

function liste(valeur: string | undefined): string[] {
  return (valeur ?? "")
    .split(",")
    .map((v) => v.trim())
    .filter(Boolean);
}

export function chargerConfig() {
  const modeLivraison = lire("MODE_LIVRAISON", "brouillon");
  if (modeLivraison !== "brouillon" && modeLivraison !== "envoi") {
    throw new Error('MODE_LIVRAISON doit valoir "brouillon" ou "envoi"');
  }

  return {
    // --- Microsoft 365 (secrets) ---
    azureTenantId: lire("AZURE_TENANT_ID"),
    azureClientId: lire("AZURE_CLIENT_ID"),
    azureClientSecret: lire("AZURE_CLIENT_SECRET"),

    // --- Boîte mail à lire ---
    /** Adresse de la boîte qui reçoit les fichiers, ex. prenom.nom@entreprise.fr */
    boiteMail: lire("BOITE_MAIL"),
    /** Expéditeurs acceptés (adresses séparées par des virgules). Vide = tous. */
    filtreExpediteurs: liste(lireOptionnel("FILTRE_EXPEDITEURS")).map((e) => e.toLowerCase()),
    /** Mot ou expression qui doit apparaître dans l'objet. Vide = aucun filtre. */
    filtreObjet: lireOptionnel("FILTRE_OBJET"),
    /** Nombre de jours en arrière à parcourir (7 = depuis le vendredi précédent). */
    joursEnArriere: Number(lire("JOURS_EN_ARRIERE", "7")),

    // --- Livraison ---
    /** Dossier OneDrive de dépôt, ex. "Rapports/Synthese hebdo". Vide = pas de dépôt OneDrive. */
    dossierOneDrive: lireOptionnel("DOSSIER_ONEDRIVE"),
    /** Destinataires de l'e-mail (séparés par des virgules). */
    destinataires: liste(lire("DESTINATAIRES")),
    /** "brouillon" : brouillon à relire dans Outlook. "envoi" : envoi automatique. */
    modeLivraison: modeLivraison as "brouillon" | "envoi",

    // --- Agent Claude (IDs écrits par `ant apply` dans claude-lock.json) ---
    agentId: lire("AGENT_ID"),
    environmentId: lire("ENVIRONMENT_ID"),
    /** Plafond de dépense par exécution, en centimes de dollar ("200" = 2 $). */
    budgetCentimes: lire("BUDGET_CENTIMES", "200"),
  };
}

export type Config = ReturnType<typeof chargerConfig>;
