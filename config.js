/* Réglages du planning Sibella. Aucun secret ici : l'identifiant d'application et celui de l'annuaire sont publics par nature. */
window.PLANNING_CONFIG = {
  clientId: '98e6b40a-63c5-43e5-8d22-112045e48002',
  tenantId: '1172a3c8-1a36-465b-931f-ed88c3debe1c',
  // Doit être identique, au caractère près, à l'URI de redirection inscrit dans Entra (plateforme « application monopage »).
  redirectUri: 'https://cabinet-sibella.github.io/planning/',
  // Équipe Teams « CABINET SIBELLA » : le planning est rangé dans la bibliothèque Documents de cette équipe.
  groupId: '461a97cc-82c0-478c-a8a8-e7c98250fa19',
  dossier: 'Planning',
  fichierDonnees: 'planning.json',
  // Index des références du registre : dossier SÉPARÉ, à réserver (droits SharePoint) aux seules personnes autorisées à voir toutes les affaires.
  dossierRefs: 'Planning - Références',
  fichierRefs: 'references.json',
  // Seules ces personnes voient les boutons de modification. La vraie barrière reste le droit « modifier » du dossier dans SharePoint.
  editeurs: [
    'pierre-emile.bacciochi@cabinetsibella.com',
    'guillaume.berwanger@cabinetsibella.com'
  ],
  // Rafraîchissement des données (secondes)
  sondage: 30
};
