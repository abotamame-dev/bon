// Accès à Outlook et OneDrive via l'API Microsoft Graph (application Entra, flux
// "client credentials"). Les identifiants restent ici : l'agent Claude n'y a pas accès.
import type { Config } from "./config.js";

const GRAPH = "https://graph.microsoft.com/v1.0";
/** Au-delà, Graph refuse la pièce jointe dans une seule requête : on envoie un lien. */
const TAILLE_MAX_PIECE_JOINTE = 3 * 1024 * 1024;

export interface FichierExcel {
  nom: string;
  contenu: Buffer;
  recuLe: string;
  expediteur: string;
  objet: string;
}

export class Graph {
  private jeton?: { valeur: string; expireA: number };

  constructor(private readonly config: Config) {}

  private async obtenirJeton(): Promise<string> {
    if (this.jeton && Date.now() < this.jeton.expireA - 60_000) return this.jeton.valeur;
    const reponse = await fetch(
      `https://login.microsoftonline.com/${this.config.azureTenantId}/oauth2/v2.0/token`,
      {
        method: "POST",
        body: new URLSearchParams({
          client_id: this.config.azureClientId,
          client_secret: this.config.azureClientSecret,
          scope: "https://graph.microsoft.com/.default",
          grant_type: "client_credentials",
        }),
      },
    );
    if (!reponse.ok) {
      throw new Error(`Connexion Microsoft refusée (${reponse.status}) : ${await reponse.text()}`);
    }
    const json = (await reponse.json()) as { access_token: string; expires_in: number };
    this.jeton = { valeur: json.access_token, expireA: Date.now() + json.expires_in * 1000 };
    return json.access_token;
  }

  private async appel(chemin: string, init: RequestInit = {}): Promise<Response> {
    const url = chemin.startsWith("https://") ? chemin : `${GRAPH}${chemin}`;
    const reponse = await fetch(url, {
      ...init,
      headers: { Authorization: `Bearer ${await this.obtenirJeton()}`, ...init.headers },
    });
    if (!reponse.ok) {
      throw new Error(`Microsoft Graph ${init.method ?? "GET"} ${chemin} → ${reponse.status} : ${await reponse.text()}`);
    }
    return reponse;
  }

  private get boite(): string {
    return `/users/${encodeURIComponent(this.config.boiteMail)}`;
  }

  /** Pièces jointes Excel des e-mails reçus depuis `depuis`, filtrées par expéditeur et objet. */
  async recupererFichiersExcel(depuis: Date): Promise<FichierExcel[]> {
    const filtre = `receivedDateTime ge ${depuis.toISOString()} and hasAttachments eq true`;
    let url: string | undefined =
      `${this.boite}/messages?$filter=${encodeURIComponent(filtre)}` +
      `&$orderby=receivedDateTime asc&$select=id,subject,from,receivedDateTime&$top=50`;

    const fichiers: FichierExcel[] = [];
    while (url) {
      const page = (await (await this.appel(url)).json()) as {
        value: Array<{
          id: string;
          subject: string | null;
          receivedDateTime: string;
          from?: { emailAddress?: { address?: string } };
        }>;
        "@odata.nextLink"?: string;
      };

      for (const message of page.value) {
        const expediteur = message.from?.emailAddress?.address?.toLowerCase() ?? "";
        const objet = message.subject ?? "";
        if (this.config.filtreExpediteurs.length && !this.config.filtreExpediteurs.includes(expediteur)) continue;
        if (this.config.filtreObjet && !objet.toLowerCase().includes(this.config.filtreObjet.toLowerCase())) continue;

        const pj = (await (
          await this.appel(`${this.boite}/messages/${message.id}/attachments?$select=id,name,size`)
        ).json()) as { value: Array<{ id: string; name: string; "@odata.type": string }> };

        for (const piece of pj.value) {
          if (piece["@odata.type"] !== "#microsoft.graph.fileAttachment") continue;
          if (!/\.(xlsx|xlsm|xls|csv)$/i.test(piece.name)) continue;
          const brut = await this.appel(`${this.boite}/messages/${message.id}/attachments/${piece.id}/$value`);
          fichiers.push({
            nom: piece.name,
            contenu: Buffer.from(await brut.arrayBuffer()),
            recuLe: message.receivedDateTime,
            expediteur,
            objet,
          });
        }
      }
      url = page["@odata.nextLink"];
    }
    return fichiers;
  }

  /** Dépose un fichier dans OneDrive et renvoie son lien web. */
  async deposerOneDrive(dossier: string, nom: string, contenu: Buffer): Promise<string> {
    const chemin = [...dossier.split("/").filter(Boolean), nom].map(encodeURIComponent).join("/");
    const reponse = await this.appel(`${this.boite}/drive/root:/${chemin}:/content`, {
      method: "PUT",
      headers: { "Content-Type": "application/octet-stream" },
      body: new Uint8Array(contenu),
    });
    const item = (await reponse.json()) as { webUrl: string };
    return item.webUrl;
  }

  /** Crée un brouillon Outlook ou envoie l'e-mail, avec la présentation jointe si elle est assez petite. */
  async livrerParMail(options: {
    objet: string;
    corpsHtml: string;
    pieceJointe?: { nom: string; contenu: Buffer };
  }): Promise<void> {
    const message = {
      subject: options.objet,
      body: { contentType: "HTML", content: options.corpsHtml },
      toRecipients: this.config.destinataires.map((address) => ({ emailAddress: { address } })),
      attachments:
        options.pieceJointe && options.pieceJointe.contenu.length <= TAILLE_MAX_PIECE_JOINTE
          ? [
              {
                "@odata.type": "#microsoft.graph.fileAttachment",
                name: options.pieceJointe.nom,
                contentBytes: options.pieceJointe.contenu.toString("base64"),
              },
            ]
          : [],
    };

    if (this.config.modeLivraison === "envoi") {
      await this.appel(`${this.boite}/sendMail`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ message, saveToSentItems: true }),
      });
    } else {
      await this.appel(`${this.boite}/messages`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(message),
      });
    }
  }
}

export { TAILLE_MAX_PIECE_JOINTE };
