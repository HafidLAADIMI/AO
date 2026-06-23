# Conception initiale de la plateforme — SIS Intégré EFPA (AOOI 01/2026/DEFR/DETFP)

Document de conception technique de référence, rédigé avant la note méthodologique. Il sert de
source unique de vérité pour la rédaction de la section "Approche méthodologique" de l'offre
(structure imposée par l'Article 9.1 du RC : A. Démarche de conduite de projet, B. Stratégie ETL
et Migration, C. Architecture SaaS et Cloud, D. Interconnexion et SSO, E. Approche IA,
F. Exploitation des données) et pour la mise en œuvre réelle si le marché est attribué.

---

## 1. Décisions de stack technique

### 1.1 Runtime back-end : Node.js LTS, pas de polyglotte additionnel

Le CPS (Art.19.3.2) impose Node.js LTS pour les services asynchrones, avec possibilité de
proposer un équivalent sous réserve de justification de robustesse institutionnelle.

Décision retenue : rester sur un stack Node.js unique pour tout le code applicatif développé par
l'équipe. Deux alternatives ont été examinées et écartées :

- **.NET** : aucun besoin fonctionnel non couvert par Node.js. L'intégration Microsoft Entra ID et
  Microsoft Graph (M365 Education) dispose de SDK officiels Node.js (MSAL Node, Microsoft Graph
  SDK for JavaScript). Ajouter .NET créerait un troisième runtime à opérer, sans gain technique
  identifié, avec un risque d'exécution réel compte tenu de l'équipe minimale de 4 profils notée
  dans la grille du RC (aucun profil .NET n'y figure) et du délai global de 7 mois.
- **FastAPI / Python comme stack de développement** : écarté pour le code métier. Conservé
  uniquement comme moteur d'inférence packagé (voir 1.3), pas comme techno développée par l'équipe.

### 1.2 Framework applicatif : NestJS plutôt qu'Express

Justification :

- Le système de modules NestJS correspond directement à l'exigence CPS Art.19.3.1 ("système
  conçu par blocs fonctionnels indépendants").
- TypeScript par défaut, sans justification supplémentaire à apporter : sûreté de typage sur un
  projet à 58→100 tenants, code auto-documenté utile pour la réversibilité totale exigée
  (Art.9.4 RC, Art.29 CPS).
- Guards/interceptors permettent un point d'audit unique appliqué uniformément sur tous les
  micro-services, répondant à la journalisation immuable (Art.19.3.5 CPS).
- `@nestjs/terminus` pour les health checks, preuve concrète du SLA 99,95% (Art.19.4.3 CPS).
- Compromis assumé : structure plus verbeuse qu'Express. Accepté car la criticité du projet
  (audit de sécurité, transfert de compétences, multi-établissements) justifie la rigueur
  structurelle plutôt que la légèreté.

### 1.3 Inférence LLM (Bot APC) : appliance externe, pas du code applicatif

Aucun moteur d'inférence LLM GPU mature n'existe en pur Node.js. Décision : déployer un moteur
d'inférence packagé (vLLM ou TGI) comme appliance conteneurisée sur l'infrastructure GPU
NVIDIA A10 exigée (Art.19.4.2 CPS), exposé en API HTTP compatible OpenAI. Ce moteur n'est pas du
code écrit par l'équipe, au même titre que PostgreSQL n'est pas écrit par l'équipe : c'est un
composant d'infrastructure consommé, pas un choix de stack de développement.

L'orchestration RAG elle-même (retrieval, construction du prompt, chaînage) reste en
**LangChain.js**, conformément au livrable nommé explicitement par le CPS (Art.21.2 : "Dossier
d'Orchestration RAG (Bot APC) : Remise du code d'orchestration LangChain.js"). Ce point est
contractuel, pas une préférence technique : y déroger expose à une non-conformité sur le
Critère A de notation (méthodologie).

### 1.4 Stockage et données

- PostgreSQL avec extension **pgvector** dans la même instance (conforme CPS, pas de vector
  store séparé).
- ORM **Prisma** (imposé CPS Art.19.3.2).
- **Supabase ou équivalent auto-hébergé** pour l'authentification bas niveau, le temps réel et le
  stockage objet (imposé CPS Art.19.3.2).

### 1.5 Multi-tenant : schéma partagé avec Row-Level Security

Décision : PostgreSQL avec RLS, `efpa_id` comme discriminant de tenant, plutôt qu'une base de
données par établissement.

Justification : la DEFR a besoin d'agrégats inter-EFPA (pilotage national, BI), ce qu'une base
par tenant rendrait coûteux à requêter. RLS sur schéma partagé est le compromis standard à
l'échelle de 58 établissements extensible à 100.

### 1.6 Traitement asynchrone : BullMQ + Redis

Tout traitement potentiellement bloquant pour l'event loop (certification de diplômes avec Barid
eSign, génération de QR code, envoi de notifications en masse après délibération) passe par une
file BullMQ adossée à Redis. La requête API répond immédiatement, un worker dédié traite le job en
arrière-plan. Permet aussi de scaler les workers de certification indépendamment des réplicas API
(Art.19.4.1 CPS, scalabilité horizontale automatique).

### 1.7 Séparation OLTP / OLAP

Le reporting (Headless BI, Art.19.5.8 CPS) n'interroge jamais directement les bases
transactionnelles. Alimentation par CDC (Change Data Capture) vers une couche analytique
(Cube.dev comme couche sémantique). Évite qu'une requête analytique lourde dégrade le service aux
apprenants en période de pointe (inscriptions, résultats de concours).

---

## 2. Découpage en services

Principe : regroupement par cohésion métier réelle (entités pivot partagées), pas par
correspondance 1:1 avec les 11 blocs fonctionnels du CPS (Art.19.5). Avec une équipe minimale de
4 profils sur 7 mois, un découpage trop fin en micro-services serait un risque d'exécution, pas
une qualité d'architecture.

### 2.1 API Gateway
Point d'entrée unique. Validation des tokens Entra ID, résolution du contexte tenant/rôle
(`efpa_id`, rôle DEFR/EFPA/enseignant/stagiaire), routage, rate limiting, intercepteur d'audit
centralisé. Ne contient aucune logique métier.

Les trois espaces extranet (Stagiaire, Enseignant/Référent, Administration — Art.19.5.7 CPS) sont
des vues SPA qui consomment les autres services via ce Gateway, pas des services backend séparés :
ils ne possèdent aucune donnée propre.

### 2.2 Référentiel Service
Données maîtres : EFPA (infrastructures, capacités), filières, modules, volumes horaires,
ressources pédagogiques, enseignants, groupes de stagiaires. Référencé par tous les autres
services (`efpa_id`, `filiere_id`, `module_id` comme clés étrangères).

### 2.3 Parcours Apprenant Service
Regroupe volontairement : admissions et concours, scolarité (inscriptions, assiduité, discipline,
abandons, changements de filière), évaluation et diplomation. Justification du regroupement : ces
trois blocs du CPS opèrent sur la même entité pivot, le dossier de l'apprenant, du dépôt de
candidature à la diplomation. Les séparer créerait des transactions distribuées permanentes sur la
même donnée. Candidat naturel à scinder plus tard si l'équipe grandit, pas à la livraison.

### 2.4 Planning & Ressources Service
Emplois du temps annuels, occupation salles/enseignants, feuilles de service, heures de vacation.
Séparé du précédent car il porte une logique de résolution de contraintes (conflits
salle/enseignant/créneau), différente d'un CRUD de dossier.

### 2.5 Stages & Insertion Service
Banque d'entreprises, affectation des stagiaires, suivi de l'insertion professionnelle.

### 2.6 Certification Service
Édition des documents officiels, certification numérique (Barid eSign), génération QR code,
archivage sécurisé. Isolé car seul service appelant une PKI externe et nécessitant le pattern
file d'attente (BullMQ) pour ne pas bloquer l'event loop.

### 2.7 BI & Reporting Service
Couche sémantique Headless BI (Cube.dev), alimentée par CDC depuis les bases OLTP. Jamais de
requête directe sur les bases transactionnelles.

### 2.8 Bot APC / IA Service
Orchestration LangChain.js, retrieval pgvector, appel au moteur d'inférence GPU externe
(vLLM/TGI). Indexe les référentiels APC, GOPM, plans de modules.

### 2.9 Interopérabilité & Migration Service
Synchronisation Moodle (groupes, notes), pont M365/Graph au-delà de l'authentification, outillage
de migration HELISA (job batch, Phase I/III, intensité dégressive après mise en service).

### 2.10 Composants d'infrastructure transverses (pas des services métier)
- Redis : files BullMQ, cache de session
- Audit Trail Store : table append-only séparée des bases transactionnelles, partitionnée
  mensuellement, conservation 12 mois glissants (Art.19.3.5 / 19.4.4 CPS)
- PostgreSQL + RLS + pgvector (voir 1.4, 1.5)

### 2.11 Communication inter-services
- Synchrone REST/JSON pour les lectures de référentiel
- Asynchrone par événements (BullMQ) pour tout ce qui ne doit pas bloquer une réponse
  utilisateur : certification, notifications, ré-indexation du Bot APC après mise à jour d'un
  référentiel pédagogique

---

## 3. Diagramme d'architecture globale

```
                                    ┌──────────────────────────────┐
                                    │   Microsoft Entra ID (SSO)     │
                                    │   OpenID Connect + MFA         │
                                    └───────────────┬────────────────┘
                                                     │ tokens
            ┌──────────────────┐   ┌──────────────────┐   ┌──────────────────┐
            │ Espace Stagiaire │   │Espace Enseignant │   │Espace Administr. │
            │      (SPA)        │   │      (SPA)        │   │      (SPA)        │
            └─────────┬─────────┘   └─────────┬─────────┘   └─────────┬─────────┘
                       └────────────────────────┼────────────────────────┘
                                                 ▼
                              ┌───────────────────────────────────┐
                              │           API GATEWAY               │
                              │ auth guard · audit interceptor      │
                              │ routing · rate limiting             │
                              └──────────────────┬───────────────────┘
       ┌──────────────┬───────────────┬──────────┼───────────┬───────────────┬─────────────┐
       ▼              ▼               ▼          ▼           ▼               ▼             ▼
┌─────────────┐┌──────────────┐┌─────────────┐┌──────────┐┌─────────────┐┌───────────┐┌────────────┐
│ Référentiel ││   Parcours   ││  Planning & ││ Stages & ││Certification││ BI & Repo. ││ Bot APC/IA │
│   Service   ││  Apprenant   ││  Ressources ││Insertion ││   Service   ││(Cube.dev)  ││(LangChain.js)│
└──────┬──────┘└──────┬───────┘└──────┬──────┘└────┬─────┘└──────┬──────┘└─────┬──────┘└─────┬──────┘
       │              │               │             │             │             ▲              │
       │              │               │             │             ▼             │ CDC          │ HTTP
       │              │               │             │      ┌─────────────┐      │              ▼
       │              │               │             │      │ BullMQ/Redis│      │      ┌───────────────┐
       │              │               │             │      └──────┬──────┘      │      │ Moteur GPU     │
       │              │               │             │             ▼             │      │ vLLM/TGI       │
       │              │               │             │      ┌─────────────┐      │      │ NVIDIA A10     │
       │              │               │             │      │ Barid eSign │      │      └───────────────┘
       │              │               │             │      │  (externe)  │      │
       │              │               │             │      └─────────────┘      │
       ▼              ▼               ▼             ▼                           │
┌───────────────────────────────────────────────────────────────────┐          │
│       PostgreSQL — RLS multi-tenant (efpa_id) + extension pgvector    │◄─────────┘
└───────────────────────────────────────────────────────────────────┘
                              ▲
                              │ sync (Phase III) / migration ETL (Phase I)
                    ┌───────────────────────┐
                    │ Interopérabilité &      │────► LMS Moodle (externe)
                    │     Migration           │────► M365 Education / Graph
                    │ (Moodle, M365, HELISA)  │
                    └───────────────────────┘

┌──────────────────────────────────────────────────────────────────────┐
│  Audit Trail Store — append-only, partitionné mensuel, 12 mois glissants│
│  alimenté par l'intercepteur du Gateway et par chaque service           │
└──────────────────────────────────────────────────────────────────────┘
```

---

## 4. Flux détaillé — Admissions et Concours

### 4.1 Diagramme de cas d'utilisation

```
                     ┌────────────────────────────────────────────────────┐
                     │             Système Admissions & Concours            │
                     │                                                      │
  Candidat ──────────┼──►( Soumettre candidature )                         │
  Candidat ──────────┼──►( Consulter statut du dossier )                   │
                     │                                                      │
  Agent EFPA ────────┼──►( Présélectionner les dossiers )                  │
  Agent EFPA ────────┼──►( Planifier une séance de concours )              │
                     │                  ▲                                   │
                     │                  │ <<include>>                      │
                     │           ( Convoquer les candidats )                │
                     │                                                      │
  Agent EFPA ────────┼──►( Saisir les notes de concours )                  │
                     │                                                      │
  Comité de          │                                                      │
  délibération ──────┼──►( Délibérer )                                     │
                     │         │                                            │
                     │         │ <<include>>                                │
                     │  ( Publier les résultats )──►( Gérer liste d'attente)│
                     │                                                      │
  Comité ────────────┼──►( Confirmer inscription définitive )              │
                     │                                                      │
  Administrateur     │                                                      │
  DEFR ───────────────┼──►( Paramétrer capacités de concours par filière ) │
                     └────────────────────────────────────────────────────┘
```

Justification : le candidat dispose uniquement de cas d'utilisation en lecture/dépôt
(soumettre, consulter) ; toute la logique de sélection, convocation, notation et délibération
reste côté agent/comité, conformément à la distinction "sélection automatique/manuelle" du CPS
qui ne donne jamais au candidat de droit d'écriture sur le processus de notation.

### 4.2 Diagramme de séquence — Délibération et publication des résultats

```
Agent EFPA       Gateway          Parcours Apprenant Svc      Référentiel Svc      Audit Store      File (BullMQ)
   │                 │                      │                       │                  │                 │
   │ POST /concours/{id}/deliberation        │                       │                  │                 │
   ├────────────────►│                      │                       │                  │                 │
   │                 │ valider token + rôle │                       │                  │                 │
   │                 ├─────────────────────►│                       │                  │                 │
   │                 │                      │ GET capacité filière  │                  │                 │
   │                 │                      ├──────────────────────►│                  │                 │
   │                 │                      │◄──────────────────────┤                  │                 │
   │                 │                      │ calcule classement,   │                  │                 │
   │                 │                      │ détermine admis/attente│                  │                 │
   │                 │                      │ persiste résultats     │                  │                 │
   │                 │                      ├────────────────────────────────────────────►                 │
   │                 │                      │  (action, agent, horodatage, avant/après)  │                 │
   │                 │                      │ enqueue notifications candidats            │                 │
   │                 │                      ├──────────────────────────────────────────────────────────────►│
   │                 │◄─────────────────────┤ 200 OK (résultats)     │                  │                 │
   │◄────────────────┤                      │                       │                  │                 │
   │                 │                      │                       │                  │    pour chaque candidat
   │                 │                      │                       │                  │    envoyer email/SMS
   │                 │                      │                       │                  │◄────────────────┤
```

Justification : la notification aux candidats est mise en file plutôt qu'envoyée en synchrone
dans la même requête, pour que la publication des résultats ne soit jamais ralentie par la
latence d'un envoi en masse. L'écriture dans l'Audit Store est synchrone et bloquante avant le
retour 200 OK : aucun résultat n'est publié sans que sa trace d'audit soit garantie persistée.

### 4.3 Diagramme de classes

```
┌─────────────────────┐            ┌──────────────────────────────┐
│      Candidat         │           │          Candidature           │
├─────────────────────┤   1     *  ├──────────────────────────────┤
│ id : UUID             │◄──────────┤ id : UUID                      │
│ nom, prenom            │           │ candidatId : UUID (FK)         │
│ cin / passeport        │           │ filiereId : UUID (FK réf.)     │
│ email, telephone       │           │ efpaId : UUID (FK réf.)        │
│ dateNaissance           │           │ statut : enum                  │
└─────────────────────┘            │ dateDepot                       │
                                     │ documentsJustificatifs[]        │
                                     └───────────────┬──────────────────┘
                                                      │ 1
                                                      │ *
                                     ┌───────────────────────────────┐
                                     │        SessionConcours           │
                                     ├───────────────────────────────┤
                                     │ id : UUID                       │
                                     │ filiereId : UUID (FK réf.)       │
                                     │ efpaId : UUID (FK réf.)          │
                                     │ dateConcours                     │
                                     │ capaciteMax                      │
                                     │ statut : enum                    │
                                     └───────────────┬───────────────────┘
                                                      │ 1
                                       ┌──────────────┴───────────────┐
                                       │ 1                             │ 1
                              ┌──────────────────┐           ┌──────────────────┐
                              │   Convocation       │           │       Note         │
                              ├──────────────────┤           ├──────────────────┤
                              │ candidatureId        │           │ candidatureId      │
                              │ dateConvocation       │           │ valeur              │
                              │ lieu, heure            │           │ matiere             │
                              │ statutEnvoi            │           │ saisiPar            │
                              └──────────────────┘           └─────────┬──────────┘
                                                                         │ *
                                                                         │ 1
                                                            ┌───────────────────────┐
                                                            │      Deliberation        │
                                                            ├───────────────────────┤
                                                            │ sessionConcoursId          │
                                                            │ dateDeliberation            │
                                                            │ comiteMembres[]             │
                                                            │ decision : enum              │
                                                            └───────────┬─────────────────┘
                                                                        │
                                                       ┌────────────────┴────────────────┐
                                                       │ 1                                │ 1
                                              ┌──────────────────┐            ┌──────────────────────┐
                                              │  ListeAttente       │            │ InscriptionDefinitive  │
                                              ├──────────────────┤            ├──────────────────────┤
                                              │ candidatureId         │            │ candidatureId           │
                                              │ rang                   │            │ dateInscription          │
                                              └──────────────────┘            │ anneeScolaire             │
                                                                                └──────────────────────┘
```

Justification : `Note` et `Convocation` référencent `Candidature` (pas directement `Candidat`),
pour permettre à un même candidat de postuler à plusieurs sessions sans collision de données.
`Deliberation` produit exactement une `ListeAttente` ou une `InscriptionDefinitive`, jamais les
deux pour la même candidature : contrainte d'exclusivité en base, pas une simple convention
applicative.

---

## 5. Flux restant à détailler

- Certification des diplômes (Barid eSign + QR code)
- Bot APC (indexation pgvector, appel au moteur d'inférence, fenêtre de contexte)
- Migration ETL des 5 ans d'historique HELISA

À compléter dans ce même fichier au fur et à mesure de la conception, avant la rédaction de la
note méthodologique finale.
