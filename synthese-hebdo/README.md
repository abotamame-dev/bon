# Synthèse hebdo

Chaque vendredi à 7h (heure de Paris), ce projet :

1. récupère dans Outlook les fichiers Excel reçus par e-mail pendant la semaine ;
2. les confie à un agent Claude (Claude Managed Agents), qui les consolide et produit une
   présentation PowerPoint avec un fichier de contrôle, vérifiés par un correcteur automatique ;
3. dépose la présentation dans OneDrive et l'envoie par Outlook (ou la laisse en brouillon).

Vos identifiants Microsoft ne quittent pas GitHub : l'agent Claude ne voit que les fichiers
Excel, sans accès à votre messagerie ni à OneDrive.

```
GitHub Actions (vendredi 7h)
  ├─ Microsoft Graph : lit les e-mails → pièces jointes Excel
  ├─ Agent Claude : Excel → synthese-AAAA-MM-JJ.pptx + controle-AAAA-MM-JJ.md
  └─ Microsoft Graph : dépôt OneDrive + e-mail (brouillon ou envoi)
```

| Fichier | Rôle |
|---|---|
| `agents/synthese-hebdo.md` | Définition de l'agent : modèle, outils, consignes |
| `environments/synthese-hebdo.yaml` | Machine isolée de l'agent (réseau limité) |
| `src/graph.ts` | Lecture d'Outlook, dépôt OneDrive, envoi d'e-mail |
| `src/agent.ts` | Lancement de l'agent, grille de contrôle, récupération des fichiers |
| `src/run.ts` | Enchaînement complet |
| `../.github/workflows/synthese-hebdo.yml` | Planification du vendredi |

## Mise en place

### 1. Application Microsoft Entra (vous ou votre service informatique)

Dans le portail Azure : **Microsoft Entra ID → Inscriptions d'applications → Nouvelle inscription**.

- **Permissions d'application** Microsoft Graph (avec consentement administrateur) :
  - `Mail.ReadWrite` : lire les e-mails et créer les brouillons ;
  - `Mail.Send` : seulement si `MODE_LIVRAISON=envoi` ;
  - `Files.ReadWrite.All` : seulement si vous utilisez `DOSSIER_ONEDRIVE`.
- **Restreignez l'accès à votre seule boîte mail** (contrôle d'accès basé sur les rôles
  d'Exchange Online pour les applications). Sans cette restriction, l'application peut lire
  toutes les boîtes de l'entreprise.
- `Files.ReadWrite.All` donne accès à tous les OneDrive de l'organisation, et il est plus
  difficile à restreindre. Si votre service informatique le refuse, laissez
  `DOSSIER_ONEDRIVE` vide : la présentation sera seulement jointe à l'e-mail.
- Créez un **secret client** et notez l'ID de locataire (tenant), l'ID d'application (client)
  et le secret.

### 2. Créer l'agent Claude (une seule fois)

Il faut l'outil en ligne de commande `ant`, version 1.30.0 ou plus, connecté à votre compte
Anthropic (`ant auth login`). Depuis le dossier `synthese-hebdo/` :

```sh
ant apply --dry-run -v agents/synthese-hebdo.md environments/synthese-hebdo.yaml   # affiche le plan, ne change rien
ant apply agents/synthese-hebdo.md environments/synthese-hebdo.yaml                # crée l'agent et l'environnement
```

Les identifiants créés sont écrits dans `claude-lock.json` : committez ce fichier. Relancez
la même commande après chaque modification de l'agent.

### 3. Secrets et variables GitHub

Dans le dépôt : **Settings → Secrets and variables → Actions**.

**Secrets** (onglet *Secrets*) :

| Nom | Valeur |
|---|---|
| `ANTHROPIC_API_KEY` | Clé API Anthropic |
| `AZURE_TENANT_ID` | ID de locataire Entra |
| `AZURE_CLIENT_ID` | ID d'application Entra |
| `AZURE_CLIENT_SECRET` | Secret client Entra |

**Variables** (onglet *Variables*) :

| Nom | Exemple | À compléter |
|---|---|---|
| `AGENT_ID` | `agent_...` | depuis `claude-lock.json` |
| `ENVIRONMENT_ID` | `env_...` | depuis `claude-lock.json` |
| `BOITE_MAIL` | `prenom.nom@entreprise.fr` | boîte qui reçoit les fichiers |
| `FILTRE_EXPEDITEURS` | `a@entreprise.fr,b@entreprise.fr` | vide = tous les expéditeurs |
| `FILTRE_OBJET` | `Rapport quotidien` | vide = aucun filtre sur l'objet |
| `JOURS_EN_ARRIERE` | `7` | période couverte |
| `DOSSIER_ONEDRIVE` | `Rapports/Synthese hebdo` | vide = pas de dépôt OneDrive |
| `DESTINATAIRES` | `prenom.nom@entreprise.fr` | séparés par des virgules |
| `MODE_LIVRAISON` | `brouillon` | `brouillon` ou `envoi` |
| `BUDGET_CENTIMES` | `200` | plafond par exécution, en centimes de dollar |

### 4. Premier test

Dans GitHub : **Actions → Synthèse hebdo → Run workflow**. Le journal affiche un lien pour
suivre l'agent en direct dans la Console Anthropic. Gardez `MODE_LIVRAISON=brouillon` tant
que le résultat ne vous convient pas.

En local, si vous préférez : copiez `.env.example` en `.env`, complétez-le, puis

```sh
npm ci
node --env-file=.env --import tsx src/run.ts
```

## Ajuster le résultat

- **Contenu de la présentation et critères de contrôle** : fonction `rubrique()` dans
  `src/agent.ts`. Chaque ligne est un critère vérifié séparément par le correcteur.
- **Consignes générales de l'agent** : le texte sous l'en-tête de
  `agents/synthese-hebdo.md`, puis relancez `ant apply`.
- **Modèle PowerPoint** : déposez votre modèle `.pptx` dans le dépôt et demandez-moi de
  l'ajouter aux fichiers envoyés à l'agent.

## Coûts et données

- Claude Managed Agents est en bêta. Chaque exécution facture les tokens du modèle plus le
  temps d'exécution de la session (0,08 $/heure). `BUDGET_CENTIMES` plafonne le total.
- Les sessions, avec les fichiers Excel envoyés, restent visibles dans la Console Anthropic
  tant que vous ne les supprimez pas.
