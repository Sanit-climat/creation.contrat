const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, ShadingType, VerticalAlign,
  PageBreak, LevelFormat, HeadingLevel, Header, Footer, PageNumber,
  NumberFormat
} = require('docx');
const fs = require('fs');

// ── helpers ──────────────────────────────────────────────────────────────────

const CYAN = "1E88C7";
const DARK = "1A1A2E";
const GRAY = "F0F4F8";

function sectionHeader(text) {
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [9360],
    rows: [new TableRow({
      children: [new TableCell({
        shading: { fill: CYAN, type: ShadingType.CLEAR },
        margins: { top: 80, bottom: 80, left: 160, right: 160 },
        width: { size: 9360, type: WidthType.DXA },
        borders: { top: noBorder(), bottom: noBorder(), left: noBorder(), right: noBorder() },
        children: [new Paragraph({
          children: [new TextRun({ text, bold: true, color: "FFFFFF", size: 22, font: "Arial" })]
        })]
      })]
    })]
  });
}

function noBorder() { return { style: BorderStyle.NONE, size: 0, color: "FFFFFF" }; }

function h1(text) {
  return new Paragraph({
    spacing: { before: 320, after: 160 },
    children: [new TextRun({ text, bold: true, size: 32, color: DARK, font: "Arial" })]
  });
}

function h2(text) {
  return new Paragraph({
    spacing: { before: 240, after: 120 },
    children: [new TextRun({ text, bold: true, size: 24, color: DARK, font: "Arial" })]
  });
}

function para(text, opts = {}) {
  return new Paragraph({
    alignment: opts.center ? AlignmentType.CENTER : AlignmentType.JUSTIFIED,
    spacing: { before: 80, after: 80 },
    children: [new TextRun({
      text,
      size: opts.size || 20,
      bold: opts.bold || false,
      italics: opts.italic || false,
      font: "Arial",
      color: opts.color || DARK
    })]
  });
}

function dotLine(label, value, boldValue = false) {
  const dots = "...........................................................................................................................";
  return new Paragraph({
    spacing: { before: 60, after: 60 },
    children: [
      new TextRun({ text: dots.substring(0, 60), font: "Arial", size: 20 }),
      new TextRun({ text: value || "", bold: boldValue, font: "Arial", size: 20, color: DARK }),
    ]
  });
}

function dotLineLabeled(label, value, boldValue = false) {
  const dots = ".....................................................";
  const labelPad = label.padEnd(0);
  return new Paragraph({
    spacing: { before: 60, after: 60 },
    children: [
      new TextRun({ text: labelPad + " ", font: "Arial", size: 20 }),
      new TextRun({ text: dots.substring(0, Math.max(5, 50 - label.length)), font: "Arial", size: 20 }),
      new TextRun({ text: " " + (value || ""), bold: boldValue, font: "Arial", size: 20, color: DARK }),
    ]
  });
}

function bullet(text) {
  return new Paragraph({
    spacing: { before: 40, after: 40 },
    indent: { left: 720, hanging: 360 },
    children: [new TextRun({ text: "\u2022  " + text, font: "Arial", size: 20 })]
  });
}

function pageBreak() {
  return new Paragraph({ children: [new PageBreak()] });
}

function emptyLine(count = 1) {
  return Array.from({ length: count }, () => new Paragraph({ children: [new TextRun("")] }));
}

function priceBox(lines) {
  return new Table({
    width: { size: 7000, type: WidthType.DXA },
    columnWidths: [7000],
    rows: [new TableRow({
      children: [new TableCell({
        shading: { fill: GRAY, type: ShadingType.CLEAR },
        margins: { top: 120, bottom: 120, left: 200, right: 200 },
        width: { size: 7000, type: WidthType.DXA },
        borders: {
          top: { style: BorderStyle.SINGLE, size: 4, color: CYAN },
          bottom: { style: BorderStyle.SINGLE, size: 4, color: CYAN },
          left: { style: BorderStyle.SINGLE, size: 4, color: CYAN },
          right: { style: BorderStyle.SINGLE, size: 4, color: CYAN }
        },
        children: lines.map(l => new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: l.text, bold: l.bold || false, size: l.size || 22, font: "Arial" })]
        }))
      })]
    })]
  });
}

// ── tva calc ─────────────────────────────────────────────────────────────────
function calcTTC(ht, tva) {
  const r = parseFloat(ht) * (1 + parseFloat(tva) / 100);
  return r.toFixed(2).replace('.', ',');
}

function formatMoney(val) {
  return parseFloat(val).toFixed(2).replace('.', ',');
}

// ── MAIN GENERATOR ───────────────────────────────────────────────────────────
function generateContract(data) {
  const d = data;
  const tvaNum = d.tva || (d.typeContrat === 'climatisation' ? '20' : '10');
  const ttcTotal = d.prestations && d.prestations.length > 0
    ? d.prestations.reduce((sum, p) => sum + parseFloat(p.prixHT || 0) * parseInt(p.quantite || 1), 0)
    : parseFloat(d.montantHT || 0);
  const ttcStr = calcTTC(ttcTotal.toFixed(2), tvaNum);
  const htStr = formatMoney(ttcTotal);

  // Label du type de contrat
  const typeLabels = {
    'maintenance': 'Contrat maintenance',
    'chaudiere': 'Contrat maintenance chaudière',
    'vmc': 'Contrat maintenance VMC',
    'thermodynamique': 'Contrat ballon thermodynamique',
    'climatisation': 'Contrat génie climatique',
    'multi': 'Contrat maintenance'
  };
  const typeLabel = typeLabels[d.typeContrat] || 'Contrat maintenance';

  // Attendu
  const equipLabel = {
    'chaudiere': 'des chaudières',
    'vmc': 'des moteurs VMC',
    'thermodynamique': 'du ballon thermodynamique',
    'climatisation': 'des installations de génie climatique',
    'multi': 'des équipements (VMC, chaudières, ramonage)',
    'maintenance': 'des équipements de maintenance'
  };

  const children = [];

  // ── PAGE DE GARDE ──────────────────────────────────────────────────────────
  children.push(
    ...emptyLine(2),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 400 },
      children: [new TextRun({ text: typeLabel, bold: true, size: 48, color: CYAN, font: "Arial" })]
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: "Maîtriser pour mieux préserver !", italics: true, size: 28, font: "Arial", color: "555555" })]
    }),
    ...emptyLine(4),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: (d.clientNom || "").toUpperCase(), bold: true, size: 56, color: "CCCCCC", font: "Arial" })]
    }),
    ...emptyLine(2),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: (d.residenceNom || d.clientAdresse1 || "").toUpperCase(), bold: true, size: 32, color: "BBBBBB", font: "Arial" })]
    }),
    ...emptyLine(6),
    pageBreak()
  );

  // ── ART 1 - DESIGNATION DES PARTIES ───────────────────────────────────────
  children.push(
    h1("ARTICLE 1 - Désignation des parties"),
    ...emptyLine(1),
    sectionHeader("ENTRE"),
    ...emptyLine(1),
    para("Les soussignées,"),
    ...emptyLine(1)
  );

  // Client selon le type (SDC/Bailleur ou Particulier)
  if (d.clientType === 'bailleur') {
    children.push(
      dotLineLabeled("", d.clientNom || "", true),
      dotLineLabeled("Représentée par", d.clientRepresentant || "", true),
      dotLineLabeled("Qui fait élection de domicile à", d.clientAdresse1 || "", false),
      dotLineLabeled("", d.clientAdresse2 || "", false),
      dotLineLabeled("", d.clientAdresse3 || "", false),
      dotLineLabeled("", "DU CLIENT", false),
    );
  } else {
    children.push(
      dotLineLabeled("Mr/Mme", d.clientNom || "", true),
      dotLineLabeled("Qui fait élection de domicile à", d.clientAdresse1 || "", false),
      dotLineLabeled("", d.clientAdresse2 || "", false),
      dotLineLabeled("", d.clientAdresse3 || "", false),
    );
  }

  children.push(
    ...emptyLine(1),
    para('Ci-après dénommée                                                                                          « le client »'),
    new Paragraph({ children: [new TextRun({ text: "D'UNE PART,", bold: true, font: "Arial", size: 20 })] }),
    ...emptyLine(1),
    sectionHeader("ET"),
    ...emptyLine(1),
    dotLineLabeled("La Société", "SANIT CLIMAT SAS au capital de 15 000 €uros", true),
    dotLineLabeled("Inscrite au Registre du Commerce des Sociétés sous le n°", "802 151 340 Marseille", true),
    dotLineLabeled("Dont le siège social est situé à", "280 avenue de Saint Antoine", true),
    dotLineLabeled("", "13015 MARSEILLE", true),
    ...emptyLine(1),
    dotLineLabeled("Représentée par", "Monsieur Rémy GALEA", true),
    dotLineLabeled("En qualité de", "Président", true),
    para('Ci-après dénommée                                                                                          « le prestataire »'),
    new Paragraph({ children: [new TextRun({ text: "DE SECONDE PART,", bold: true, font: "Arial", size: 20 })] }),
    ...emptyLine(1),
    para('Les soussignées désignés ensemble par                                                         « les parties »'),
    ...emptyLine(1),
    para(`Attendu que ${d.clientRepresentant || d.clientNom || ""} (le client) ${d.clientType === 'bailleur' ? `représenté par ${d.clientRepresentant2 || ""}` : ""} souhaite souscrire un contrat d'entretien et de maintenance ${equipLabel[d.typeContrat] || "des équipements"} de la résidence ${d.residenceNom || ""} située ${d.residenceAdresse || ""}.`, { bold: false }),
    ...emptyLine(1),
    para("Attendu que la société SANIT CLIMAT (le prestataire) déclare disposer des compétences et des moyens nécessaires pour l'exécution des prestations de service attendus par « le client » et consignées dans le contrat et ses annexes."),
    ...emptyLine(1),
    new Paragraph({ children: [new TextRun({ text: "IL EST CONVENU ET ARRETE CE QUI SUIT.", bold: true, font: "Arial", size: 20 })] }),
    pageBreak()
  );

  // ── ART 2 - DEFINITION ────────────────────────────────────────────────────
  children.push(
    h1("ARTICLE 2 - Définition"),
    ...emptyLine(1),
    para("Par le présent contrat, pour la durée de celui-ci et aux conditions énoncées ci-après, le client confie au prestataire qui accepte, le soin d'assurer en exclusivité l'entretien des installations décrites en annexe, dont il est et reste propriétaire."),
    ...emptyLine(1),
    para("Cette prestation de service implique, pour le prestataire, une obligation de moyens."),
    ...emptyLine(1),
    sectionHeader("DOCUMENTS CONTRACTUELS"),
    ...emptyLine(1),
    para("Le présent marché est constitué par les documents énumérés ci-après par ordre de priorité décroissante."),
    ...emptyLine(1),
    para("Les pièces particulières", { bold: true }),
    bullet("Le présent contrat de maintenance,"),
    bullet("Les annexes au contrat de maintenance :"),
    new Paragraph({ indent: { left: 1080 }, spacing: { before: 40, after: 40 }, children: [new TextRun({ text: "\u2022  Inventaire des équipements,", font: "Arial", size: 20 })] }),
    new Paragraph({ indent: { left: 1080 }, spacing: { before: 40, after: 40 }, children: [new TextRun({ text: "\u2022  Le tarif horaire avec les majorations applicables et le forfait de déplacement,", font: "Arial", size: 20 })] }),
    new Paragraph({ indent: { left: 1080 }, spacing: { before: 40, after: 40 }, children: [new TextRun({ text: "\u2022  La définition des interventions programmées,", font: "Arial", size: 20 })] }),
    new Paragraph({ indent: { left: 1080 }, spacing: { before: 40, after: 40 }, children: [new TextRun({ text: "\u2022  Dossier des ouvrages exécutés des installations.", font: "Arial", size: 20 })] }),
    ...emptyLine(1),
    para("Les pièces générales", { bold: true }),
    bullet("Les documents techniques unifiés (D.T.U.) en vigueur,"),
    bullet("Les normes en vigueur,"),
    bullet("Les réglementations en vigueur concernant la législation du travail,"),
    bullet("Le règlement intérieur appliqué pour les bâtiments du client."),
    ...emptyLine(1),
    para("Les documents applicables sont ceux en vigueur au premier jour du mois de l'établissement du présent contrat."),
    pageBreak()
  );

  // ── ART 3 - OBLIGATIONS DU PRESTATAIRE ────────────────────────────────────
  children.push(
    h1("ARTICLE 3 - Obligations du prestataire"),
    h2("NATURE DES PRESTATIONS ET EXCLUSIONS"),
    ...emptyLine(1),
    sectionHeader("PRESTATIONS D'ENTRETIEN COURANT"),
    ...emptyLine(1),
    para("Le service d'entretien courant à la charge du prestataire comprendra :"),
    ...emptyLine(1),
    para("L'intervention du personnel qualifié disposant de l'outillage nécessaire à l'ensemble des opérations décrites dans les annexes."),
    para("Le contrôle par un responsable qualifié du bon fonctionnement général des équipements pris en charge et de l'exécution de la prestation prévue au contrat."),
    para("Le prestataire devra pouvoir, sur simple demande, justifier auprès du client de l'exécution de ses prestations."),
    para("Périodiquement, à l'échéance anniversaire, le responsable établira un rapport sur la période écoulée, traitant des principales observations formulées, de leur mise en œuvre et des préconisations du prestataire sur les gros travaux à envisager pour assurer le maintien des conditions normales de fonctionnement du matériel."),
    ...emptyLine(1),
    para("3.1 Entretien forfaitaire P2", { bold: true }),
    para("Cet entretien couvre l'ensemble des appareils énumérés en pièce jointe, et comprend :"),
    bullet("1 visite annuelle systématique et préventive dont le détail est défini sur la nomenclature des opérations d'entretien en annexe ;"),
    bullet("La prise en charge par le prestataire des frais de main-d'œuvre consécutifs aux dépannages demandés expressément par le client ou les utilisateurs."),
    ...emptyLine(1),
  );

  // P2+P3 pour chaudières avec garantie pièces
  if (d.garantiePieces === 'oui') {
    children.push(
      para("3.2 Entretien forfaitaire et garantie pièces P2 + P3", { bold: true }),
      para("Cet entretien et garantie pièces ne couvrent que ceux des appareils expressivement visés par ce type d'entretien en Annexe. Il comprend exclusivement :"),
      bullet("Les prestations prévues pour l'entretien forfaitaire P2 répertoriées à l'alinéa 3.1 ;"),
      bullet("La prise en charge par le prestataire des frais de réparation ou de remplacement des pièces usagées, excepté les pièces émaillées ou peintes, les châssis, les scellements des appareils et le remplacement complet d'un appareil ;"),
      bullet("Le détartrage de l'échangeur sanitaire, à l'appréciation du prestataire, à raison d'un par an maximum."),
      ...emptyLine(1),
    );
  }

  children.push(
    sectionHeader("LIMITE DE LA PRESTATION"),
    ...emptyLine(1),
    para("Le service d'entretien est expressément limité aux matériels et aux opérations précisées dans les annexes. Toute modification de la prestation doit faire l'objet d'un avenant."),
    para("L'entretien n'est prévu que pour les installations fixes. Les prestations pour les appareils mobiles, alimentés par des canalisations souples, des cordons flexibles, par prise de courant, connecteurs, etc., ne sont pas comprises dans le présent contrat."),
    para("Ne sont pas couverts par le contrat entretien forfaitaire P2 les pièces ou les appareils de rechange. Ils seront facturés séparément au client."),
    ...emptyLine(1),
    sectionHeader("DEPANNAGES"),
    ...emptyLine(1),
    para("En dehors des visites planifiées et pendant les jours et heures œuvrées (de 08h00 à 17h00 du Lundi au vendredi hors week-ends et jours fériés), le prestataire mettra son service de dépannage à la disposition du client sous 8 (Huit) heures selon les modalités fixées aux conditions financières figurant en annexe."),
    new Paragraph({ spacing: { before: 80, after: 80 }, children: [new TextRun({ text: "Les dépannages ne sont pas inclus", bold: true, font: "Arial", size: 20 }), new TextRun({ text: " dans le montant forfaitaire annuel, ils feront l'objet d'une facturation séparée à laquelle sera joint l'attachement correspondant.", font: "Arial", size: 20 })] }),
    ...emptyLine(1),
    sectionHeader("FOURNITURES"),
    ...emptyLine(1),
    para("Pour assurer l'entretien courant, le prestataire n'aura à sa charge que les ingrédients et consommables nécessaires à sa prestation (huile, graisse, chiffons, petite visserie et boulonnerie). Les autres fournitures et/ou pièces détachées dont l'approvisionnement sera assuré par le prestataire seront facturées au client."),
    ...emptyLine(1),
    sectionHeader("TRAVAUX HORS CONTRAT"),
    ...emptyLine(1),
    new Paragraph({ spacing: { before: 80, after: 80 }, children: [new TextRun({ text: "1)", bold: true, font: "Arial", size: 20 }), new TextRun({ text: " La fourniture du matériel et de la main d'œuvre nécessaires au remplacement de pièces usagées ou défectueuses, quelles qu'elles soient, et quelle qu'en soit la cause,", font: "Arial", size: 20 })] }),
    ...["l'équilibrage de base de l'installation,", "les travaux nécessaires à la sécurité de marche des installations et leur mise en conformité avec les règlements en vigueur,", "les travaux de réparation ou remise en bon état de marche des installations (gros entretien),", "tous travaux en parties privatives,", "et en règle générale, tous travaux non explicitement définis dans le présent contrat."].map(t => new Paragraph({ indent: { left: 720 }, spacing: { before: 40, after: 40 }, children: [new TextRun({ text: t, font: "Arial", size: 20 })] })),
    ...emptyLine(1),
    para("Ces travaux pourront être exécutés par le prestataire après accord écrit du client."),
    para("En cas de refus du client, de procéder à l'exécution des travaux mentionnés au § 1), le prestataire pourra suspendre l'exécution de ses prestations sous la seule réserve d'en aviser le client par lettre recommandée."),
    pageBreak()
  );

  // ── ART 4 - OBLIGATIONS DU CLIENT ─────────────────────────────────────────
  children.push(
    h1("ARTICLE 4 – Obligations du client"),
    ...emptyLine(1),
    sectionHeader("CONDUITE DES INSTALLATIONS"),
    ...emptyLine(1),
    para("Le client assure normalement la conduite et la surveillance de son installation et toutes les prestations qui en découlent autres que celles à la charge du prestataire."),
    ...emptyLine(1),
    sectionHeader("ENGAGEMENTS DU CLIENT"),
    ...emptyLine(1),
    para("Le client s'engage :"),
    ...emptyLine(1),
    ...["A obtenir les certificats de conformité relatifs aux installations dont font partie les appareils pris en charge par le prestataire au titre du présent contrat.",
      "A utiliser les installations dans des conditions normales de fonctionnement, à l'abri de toute malveillance ou action anormale.",
      "A fournir les combustibles, les fluides et l'électricité nécessaires, en qualité et en quantité, au fonctionnement de ses installations ainsi qu'aux travaux éventuels.",
      "A justifier de toutes assurances utiles garantissant les risques liés à l'existence et à l'exploitation de ses installations.",
      "A garantir le libre accès des lieux au prestataire.",
      "A garantir au prestataire toutes conditions de sécurité réglementaires.",
      "A fournir au prestataire les renseignements nécessaires à l'exécution de l'entretien courant, en particulier, plans, schémas, notices des constructeurs, installateurs et fournisseurs, etc.",
      "A exécuter les travaux de gros entretien préconisés par le prestataire et en particulier ceux nécessaires à la mise en conformité des installations.",
      "A prendre à sa charge les visites de contrôle par un organisme agréé si la législation l'exige."
    ].map(t => new Paragraph({ indent: { left: 360 }, spacing: { before: 60, after: 60 }, children: [new TextRun({ text: t, font: "Arial", size: 20 })] })),
    pageBreak()
  );

  // ── ART 5 - CONDITIONS FINANCIERES ────────────────────────────────────────
  children.push(
    h1("ARTICLE 5 - Conditions financières"),
    ...emptyLine(1),
    sectionHeader("REDEVANCE"),
    ...emptyLine(1),
    para("En contrepartie des prestations contractuelles d'entretien courant, le prestataire percevra une redevance annuelle forfaitaire d'un montant de :"),
    ...emptyLine(1),
  );

  // Tableau des prix
  if (d.prestations && d.prestations.length > 0) {
    // Multi-prestations
    const priceLinesTable = [];
    let total = 0;
    d.prestations.forEach(p => {
      const ht = parseFloat(p.prixHT || 0) * parseInt(p.quantite || 1);
      total += ht;
      priceLinesTable.push({ text: `${p.visites || '1'} visite(s) annuelle(s)`, bold: true });
      priceLinesTable.push({ text: `Prestation maintenance ${p.designation} (QTE ${p.quantite})`, bold: true });
      priceLinesTable.push({ text: `${formatMoney(p.prixHT)} €uros H.T. soit ${calcTTC(p.prixHT, tvaNum)} €uros T.T.C.`, bold: true, size: 24 });
    });
    children.push(priceBox(priceLinesTable));
  } else {
    // Prestation unique
    children.push(priceBox([
      { text: `${d.nombreVisites || '1'} visite(s) annuelle(s)`, bold: true },
      { text: `${htStr} €uros H.T. soit ${ttcStr} €uros T.T.C.`, bold: true, size: 28 }
    ]));
  }

  children.push(
    ...emptyLine(1),
    para(`au taux actuel de T.V.A. de ${tvaNum}%. En cas de variation du taux de T.V.A. ou de création d'une nouvelle taxe, le prix sera toujours basé sur le montant hors taxes auquel seront appliqués les nouveaux taux.`),
    ...emptyLine(1),
    sectionHeader("REVISION DES PRIX"),
    ...emptyLine(1),
    para("La redevance hors taxes fera l'objet d'une révision en fonction des variations des indices économiques. Elles seront révisées en fin d'année, selon la formule suivante :"),
    ...emptyLine(1),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: "P = P₀ [(0,80 × ICHT IME / ICHT IME₀) + (0,20 × Fsd2 / Fsd20)]", bold: true, size: 24, font: "Arial" })]
    }),
    ...emptyLine(1),
    sectionHeader("CONDITIONS DE PAIEMENT"),
    ...emptyLine(1),
    para("Les prestations forfaitaires sont payables d'avance (redevance annuelle). Les factures seront établies annuellement et adressées par le prestataire au début de chaque année pour la période à venir."),
    para("Les règlements seront effectués par virement bancaire à date de réception de chaque facture."),
    para("Les prestations effectuées hors contrat seront facturées en régie sur la base du tarif horaire et déplacement joint en annexe, et payables par le client à réception de la facture."),
    ...emptyLine(1),
    sectionHeader("DISPOSITIONS PARTICULIERES"),
    ...emptyLine(1),
    para("En cas de non-paiement des sommes dues, le contrat pourra être résilié de plein droit, aux torts et griefs du client, après mise en demeure adressée par le prestataire au client, restée infructueuse pendant un délai de 15 jours."),
    para("En outre, et sans préjudice de ce qui précède, tout retard de paiement entraînera l'exigibilité immédiate de toutes sommes dues, augmentées d'un intérêt calculé au taux d'intérêt légal majoré de cinq points, calculé par jour calendaire de retard."),
    pageBreak()
  );

  // ── ART 6 - DUREE ─────────────────────────────────────────────────────────
  children.push(
    h1("ARTICLE 6 - Durée du contrat"),
    ...emptyLine(1),
    sectionHeader("PERIODE D'EXECUTION DES PRESTATIONS"),
    ...emptyLine(1),
    new Paragraph({
      spacing: { before: 80, after: 80 },
      children: [
        new TextRun({ text: "Le présent contrat est conclu pour une période de ", font: "Arial", size: 20 }),
        new TextRun({ text: `${d.duree || "1 an"}`, bold: true, font: "Arial", size: 20 }),
        new TextRun({ text: " allant du ", font: "Arial", size: 20 }),
        new TextRun({ text: d.dateDebut || "", bold: true, font: "Arial", size: 20 }),
        new TextRun({ text: " au ", font: "Arial", size: 20 }),
        new TextRun({ text: d.dateFin || "", bold: true, font: "Arial", size: 20 }),
      ]
    }),
    ...emptyLine(1),
    para("A l'arrivée du terme, il se renouvellera d'année en année par tacite reconduction pour des périodes d'un an, sauf dénonciation par l'une ou l'autre des parties notifiées à l'autre par lettre recommandée avec avis de réception avec un préavis de 3 mois minimum avant la date anniversaire."),
    pageBreak()
  );

  // ── ART 7 - RESPONSABILITE ────────────────────────────────────────────────
  children.push(
    h1("ARTICLE 7 - Responsabilité du prestataire"),
    ...emptyLine(1),
    para("La responsabilité encourue par le prestataire vis-à-vis du client découle d'une obligation de moyens. En conséquence, sa responsabilité sera engagée en cas de faute prouvée à son égard."),
    ...emptyLine(1),
    sectionHeader("ASSURANCE"),
    ...emptyLine(1),
    para("Le prestataire déclare être titulaire d'un contrat d'assurance garantissant les conséquences pécuniaires de sa responsabilité civile susceptible d'être engagée à l'occasion de l'exécution du présent contrat."),
    ...emptyLine(1),
    sectionHeader("CAS DE FORCE MAJEURE"),
    ...emptyLine(1),
    para("Aucune partie n'est responsable vis-à-vis de l'autre de l'inexécution ou d'un retard dans l'exécution de ses obligations issues des présentes en raison d'un cas de force majeure."),
    para("Sont considérés comme cas de force majeure, outre ceux habituellement reconnus par la jurisprudence et les tribunaux français, toute décision des autorités, guerre, émeute, grève illégale ou d'envergure nationale, blocage des transports ou des réseaux électroniques, catastrophe naturelle ou encore tout autre évènement en dehors de la volonté des parties."),
    ...emptyLine(1),
    sectionHeader("PRISE EN CHARGE DES INSTALLATIONS"),
    ...emptyLine(1),
    para("Le prestataire établira dans un délai de 3 (trois) mois, à compter de la date de démarrage de la prestation, un procès-verbal contradictoire de prise en charge des installations."),
    ...emptyLine(1),
    sectionHeader("CLASSEUR DE SUIVI DE LA PRESTATION"),
    ...emptyLine(1),
    para("Dès lors que le « procès verbal contradictoire de prise en charge » des installations est validé par le client, le prestataire met en place sur site un classeur de suivi de la prestation dans lequel sont consignés :"),
    ...["le détail des opérations programmées objets du présent contrat,", "les visites systématiques de maintenance préventive avec mention de la date, de la nature des opérations,", "le double des feuillets d'attachement effectués lors des interventions de dépannage,", "les modifications et travaux effectués à l'initiative du prestataire ou sur demande du client,", "les anomalies de fonctionnement décelées et incidents majeurs."].map(t => new Paragraph({ indent: { left: 360 }, spacing: { before: 40, after: 40 }, children: [new TextRun({ text: t, font: "Arial", size: 20 })] })),
    pageBreak()
  );

  // ── ART 8 - CESSION / RESILIATION ─────────────────────────────────────────
  children.push(
    h1("ARTICLE 8 – Cession – Résiliation Anticipée"),
    ...emptyLine(1),
    para("Les présentes dispositions contractuelles sont opposables aux ayants droits à quelque titre que ce soit, locataires ou successeurs éventuels du client."),
    ...emptyLine(1),
    sectionHeader("TRANSFERT DU CONTRAT"),
    ...emptyLine(1),
    para("Le présent contrat a été conclu intuitu personae. En conséquence, les droits et obligations découlant du présent contrat ne pourront être transférés, cédés ou apportés par l'une ou l'autre des parties sans l'autorisation expresse préalable de l'autre partie."),
    ...emptyLine(1),
    sectionHeader("RESILIATION ANTICIPEE"),
    ...emptyLine(1),
    para("Si le client n'a pas procédé au règlement des redevances dans les délais stipulés au présent contrat, le prestataire pourra, huit jours après avoir effectué une mise en demeure par lettre recommandée restée sans effet, suspendre l'ensemble de ses prestations contractuelles."),
    para("Toute résiliation anticipée par le client en dehors du cas mentionné ci-dessus donne lieu au dédommagement du prestataire pour le préjudice commercial et financier subi à cause de l'interruption du contrat. Ce dédommagement est fixé à 50% du total des redevances qui auraient été versées jusqu'à la fin du contrat."),
    pageBreak()
  );

  // ── ART 9 - CONTENTIEUX ───────────────────────────────────────────────────
  children.push(
    h1("ARTICLE 9 - Contentieux"),
    ...emptyLine(1),
    sectionHeader("REGLEMENT DES LITIGES"),
    ...emptyLine(1),
    para("Tout différend portant sur la validité, l'interprétation ou l'exécution du présent contrat qui n'aura pu être résolu à l'amiable entre les parties, sera exclusivement porté devant le Tribunal de Commerce de Marseille (Bouches du Rhône)."),
    pageBreak()
  );

  // ── ART 10 - SIGNATURES ───────────────────────────────────────────────────
  children.push(
    h1("ARTICLE 10 – Signatures des parties"),
    ...emptyLine(4),
    new Table({
      width: { size: 9360, type: WidthType.DXA },
      columnWidths: [4680, 4680],
      rows: [
        new TableRow({
          children: [
            new TableCell({
              borders: { top: noBorder(), bottom: noBorder(), left: noBorder(), right: noBorder() },
              width: { size: 4680, type: WidthType.DXA },
              children: [
                new Paragraph({ children: [new TextRun({ text: "Pour le prestataire :", bold: true, font: "Arial", size: 20 })] }),
                ...emptyLine(3),
                new Paragraph({ children: [new TextRun({ text: "Date : ...............................", font: "Arial", size: 20 })] }),
                ...emptyLine(2),
                new Paragraph({ children: [new TextRun({ text: "Cachet :", font: "Arial", size: 20 })] }),
                ...emptyLine(4),
                new Paragraph({ children: [new TextRun({ text: "Signature :", font: "Arial", size: 20 })] }),
              ]
            }),
            new TableCell({
              borders: { top: noBorder(), bottom: noBorder(), left: noBorder(), right: noBorder() },
              width: { size: 4680, type: WidthType.DXA },
              children: [
                new Paragraph({ children: [new TextRun({ text: "Pour le client :", bold: true, font: "Arial", size: 20 })] }),
                ...emptyLine(3),
                new Paragraph({ children: [new TextRun({ text: "Date : ...............................", font: "Arial", size: 20 })] }),
                ...emptyLine(2),
                new Paragraph({ children: [new TextRun({ text: "Cachet :", font: "Arial", size: 20 })] }),
                ...emptyLine(4),
                new Paragraph({ children: [new TextRun({ text: "Signature :", font: "Arial", size: 20 })] }),
              ]
            }),
          ]
        })
      ]
    })
  );

  // ── BUILD DOCUMENT ────────────────────────────────────────────────────────
  const doc = new Document({
    styles: {
      default: {
        document: { run: { font: "Arial", size: 20, color: DARK } }
      }
    },
    sections: [{
      properties: {
        page: {
          size: { width: 11906, height: 16838 }, // A4
          margin: { top: 1134, right: 1134, bottom: 1134, left: 1134 }
        }
      },
      headers: {
        default: new Header({
          children: [
            new Table({
              width: { size: 9638, type: WidthType.DXA },
              columnWidths: [9638],
              rows: [new TableRow({
                children: [new TableCell({
                  shading: { fill: CYAN, type: ShadingType.CLEAR },
                  margins: { top: 60, bottom: 60, left: 160, right: 160 },
                  width: { size: 9638, type: WidthType.DXA },
                  borders: { top: noBorder(), bottom: noBorder(), left: noBorder(), right: noBorder() },
                  children: [new Paragraph({
                    children: [new TextRun({ text: "SANIT CLIMAT", bold: true, color: "FFFFFF", size: 22, font: "Arial" })]
                  })]
                })]
              })]
            }),
          ]
        })
      },
      footers: {
        default: new Footer({
          children: [
            new Paragraph({
              border: { top: { style: BorderStyle.SINGLE, size: 6, color: CYAN } },
              spacing: { before: 80 },
              children: [
                new TextRun({ text: "07.68.63.34.20    contact@sanit-climat.fr    280 avenue de Saint Antoine  13015 MARSEILLE", font: "Arial", size: 16, color: "888888" }),
              ]
            })
          ]
        })
      },
      children
    }]
  });

  return doc;
}

// ── CLI entry point ───────────────────────────────────────────────────────────
const args = process.argv.slice(2);
if (args.length < 2) {
  console.error("Usage: node generate_contract.js <data.json> <output.docx>");
  process.exit(1);
}

const data = JSON.parse(fs.readFileSync(args[0], 'utf8'));
const doc = generateContract(data);
Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync(args[1], buffer);
  console.log("OK:" + args[1]);
}).catch(e => {
  console.error("ERR:" + e.message);
  process.exit(1);
});
