# Architecture Technique - Plateforme SGG
## AOOI N°01/2026/DSS - Direction de la Stratégie et des Statistiques

---

## Pattern Architectural

**Microservices avec domaines clairs et communication asynchrone via message broker.**

Chaque service a une responsabilité unique, une base de données propre adaptée à sa nature, et communique avec les autres via RabbitMQ pour les événements asynchrones. L'API Gateway est le seul point d'entrée pour tous les clients.

```
[React SPA]
     |
     | HTTPS
     v
[API Gateway]  <-->  [Keycloak]
     |
     +---> [API Principale]  ---> [PostgreSQL]
     |            |
     |            +--> publie events --> [RabbitMQ]
     |                                       |
     +---> [Service Rapport]                 +--> [Service Audit]    --> [MongoDB]
     |     (lecture via vues PostgreSQL)     |
     |                                       +--> [Service Notification] --> [PostgreSQL]
     |
     +---> [MinIO / Garage] (documents)
```

---

## Services et Technologies

### 1. API Principale

**Rôle :** Toutes les opérations CRUD. Gestion des référentiels, saisie des données, circuit de validation, gestion des utilisateurs, tableaux de bord.

| Technologie | Version | Raison du choix |
|---|---|---|
| Kotlin | 2.4.0 | Null safety, data classes, coroutines, moins de boilerplate que Java |
| Spring Boot | 4.0.6 | Framework enterprise mature, first-class support Kotlin, Spring Security |
| Spring Data JPA | inclus dans Spring Boot 4.0.6 | ORM robuste avec Kotlin extensions |
| PostgreSQL | 18.4 | ACID complet, Row-Level Security, JSONB, partitionnement |
| TimescaleDB | extension PostgreSQL 18 | Séries temporelles pour l'historique des indicateurs 2021-2030 |
| Redis | 7.x | Cache des dashboards, sessions, pub/sub |
| Gradle (Kotlin DSL) | 8.x | Build tool natif Kotlin |

**Points clés Spring Boot 4.0 :**
- Requiert Java 17 minimum
- Support natif des coroutines Kotlin dans WebMVC et WebFlux
- API versioning stable pour les endpoints HTTP
- Modules plus petits et mieux découpés
- JSpecify pour la null-safety interopérabilité Java/Kotlin

---

### 2. Service Audit et Traçabilité

**Rôle :** Enregistrement de toutes les actions critiques de la plateforme. Qui a fait quoi, quand, depuis quelle IP. Consomme les événements de RabbitMQ de manière asynchrone.

| Technologie | Version | Raison du choix |
|---|---|---|
| Kotlin | 2.4.0 | Cohérence avec le reste des services JVM |
| Spring Boot | 4.0.6 | Même stack JVM, consommateur RabbitMQ natif |
| MongoDB | 8.3.2 | Schéma flexible pour les logs, append-only, haute fréquence d'écriture |
| Spring AMQP | inclus | Intégration RabbitMQ dans Spring |

**Pourquoi MongoDB pour l'audit :**
Les logs ont une structure variable selon le type d'action. MongoDB permet d'enregistrer un événement de saisie, un événement de validation ou un événement d'administration dans la même collection sans contrainte de schéma fixe. Les requêtes sont principalement en lecture séquentielle par date.

---

### 3. Service Notification

**Rôle :** Envoi des alertes d'échéance, notifications de validation/rejet, rappels automatiques. Consomme les événements RabbitMQ et gère les templates d'emails.

| Technologie | Version | Raison du choix |
|---|---|---|
| Kotlin | 2.4.0 | Stack JVM unifiée |
| Spring Boot | 4.0.6 | JavaMailSender, Spring Scheduler, Spring AMQP |
| Thymeleaf | 3.x | Templates HTML pour les emails |
| PostgreSQL | 18.4 | Stockage des préférences de notification et historique d'envoi |
| Spring Scheduler | inclus | Jobs planifiés pour les rappels automatiques |

---

### 4. Service Rapport

**Rôle :** Génération des rapports PDF et Excel à la demande et de façon planifiée. Accède directement aux vues PostgreSQL dédiées à la lecture analytique.

| Technologie | Version | Raison du choix |
|---|---|---|
| Python | 3.12 | Écosystème sans équivalent pour la génération de documents |
| FastAPI | 0.136.3 | Async natif, documentation OpenAPI automatique, performances |
| ReportLab / WeasyPrint | latest | Génération PDF avec templates |
| openpyxl | latest | Génération Excel avec formules et mises en forme |
| Jinja2 | latest | Templates de rapports |
| psycopg3 | latest | Driver PostgreSQL async pour Python |

**Accès aux données :**
Le service rapport se connecte à PostgreSQL avec un utilisateur dédié en lecture seule. Il accède uniquement au schéma `reporting` qui expose des vues pré-calculées. Les tables brutes du schéma `public` sont invisibles pour cet utilisateur.

```sql
CREATE USER rapport_svc WITH PASSWORD '...' NOSUPERUSER NOCREATEDB;
GRANT CONNECT ON DATABASE sgg TO rapport_svc;
GRANT USAGE ON SCHEMA reporting TO rapport_svc;
GRANT SELECT ON ALL TABLES IN SCHEMA reporting TO rapport_svc;
```

---

### 5. API Gateway

**Rôle :** Point d'entrée unique. Validation des tokens JWT, routage vers les services, rate limiting, logging des requêtes entrantes.

| Technologie | Version | Raison du choix |
|---|---|---|
| Spring Cloud Gateway | 4.x | Natif Spring, intégration Keycloak directe, réactif |
| Kotlin | 2.4.0 | Cohérence |

---

### 6. Gestion des Identités et Accès (IAM)

| Technologie | Version | Raison du choix |
|---|---|---|
| Keycloak | 26.6.3 | SSO, RBAC fin, JWT, gestion des realms, battle-tested en contexte gouvernemental |

**Configuration Keycloak pour ce projet :**
- Un realm `sgg-platform`
- Rôles : `ADMIN`, `VALIDATEUR_CENTRAL`, `VALIDATEUR_REGIONAL`, `SAISISSEUR`, `LECTEUR`, `DECIDEUR`
- Attributs utilisateur : région assignée, filières autorisées
- Token TTL : 15 minutes access token, 8 heures refresh token

**Nouveautés Keycloak 26.6 utilisées :**
- JWT Authorization Grant pour les communications inter-services
- Workflows d'authentification configurables
- Zero-downtime patch releases

---

### 7. Stockage des Documents

**Situation MinIO en 2026 :**
Le dépôt GitHub minio/minio a été archivé le 25 avril 2026. Les binaires existants continuent de fonctionner mais ne reçoivent plus de mises à jour de sécurité ni de nouvelles fonctionnalités.

**Décision :** Utiliser **Garage** comme alternative.

| Technologie | Version | Raison du choix |
|---|---|---|
| Garage | latest stable | Entièrement S3-compatible, maintenu par une organisation à but non lucratif, léger, open source Apache 2.0 |

Garage expose la même API S3 que MinIO. Tout le code qui utilisait le SDK S3 avec MinIO fonctionne sans modification avec Garage.

---

### 8. Message Broker

| Technologie | Version | Raison du choix |
|---|---|---|
| RabbitMQ | 4.3.1 | Mature, simple à opérer, parfait pour ce volume, AMQP 0.9.1 |

**Événements publiés par l'API principale :**
- `data.submitted` : une saisie a été soumise à validation
- `data.validated` : une saisie a été validée
- `data.rejected` : une saisie a été rejetée avec motif
- `deadline.approaching` : échéance de saisie dans 48h
- `deadline.missed` : échéance de saisie dépassée
- `user.action` : toute action critique d'un utilisateur (pour l'audit)

---

### 9. Frontend

| Technologie | Version | Raison du choix |
|---|---|---|
| React | 19.x | Composants riches, écosystème dashboard |
| Vite | 8.0.16 | Build Rust-based, 3x plus rapide au démarrage, pas de SSR inutile |
| TypeScript | 5.x | Typage fort, cohérence avec le backend |
| TanStack Query | latest (2026-06-02) | Gestion du server state, cache automatique, sync |
| TanStack Table | latest | Grilles de données avec filtres, tri, pagination, export |
| Apache ECharts 6.0 | 6.x | Visualisation avancée, thèmes dynamiques, nouveaux types de charts |
| React Leaflet | latest | Cartographie des 12 régions avec données superposées |
| Tailwind CSS | 4.x | Utilitaires CSS, cohérence visuelle |
| shadcn/ui | latest | Composants accessibles, pas de dépendance lourde |
| Zustand | latest | State management léger pour l'état client |

**Pourquoi Vite 8 et pas Next.js :**
La plateforme est entièrement protégée par authentification. Aucune page n'est publique, donc le SSR de Next.js n'apporte rien (ni SEO, ni First Contentful Paint amélioré). Vite 8 avec le nouveau toolchain Rust est plus simple et plus performant pour ce cas d'usage.

**Nouveautés Apache ECharts 6.0 utiles :**
- Thèmes dynamiques et changement de thème à la volée
- Chord charts pour visualiser les relations filières/régions
- Broken axes pour gérer les grandes disparités de valeurs entre régions
- Custom series réutilisables

---

## Schéma des bases de données

| Service | Base | Technologie |
|---|---|---|
| API Principale | sgg_db (schéma public + reporting) | PostgreSQL 18.4 + TimescaleDB |
| Service Audit | sgg_audit | MongoDB 8.3.2 |
| Service Notification | sgg_notifications | PostgreSQL 18.4 |
| Keycloak | keycloak_db | PostgreSQL 18.4 (instance séparée) |

---

## Infrastructure et Déploiement

| Technologie | Version | Rôle |
|---|---|---|
| Docker | latest | Containerisation de chaque service |
| Docker Compose | latest | Orchestration locale et production initiale |
| Nginx | 1.27.x | Reverse proxy, SSL termination, compression gzip, rate limiting |
| GitHub Actions / GitLab CI | - | Pipeline CI/CD : tests, build, déploiement |

**Structure Docker Compose :**
```
services:
  gateway          (Spring Cloud Gateway)
  api              (Spring Boot / Kotlin)
  audit            (Spring Boot / Kotlin)
  notification     (Spring Boot / Kotlin)
  rapport          (FastAPI / Python)
  keycloak         (Keycloak 26.6.3)
  postgres         (PostgreSQL 18.4)
  postgres-keycloak (PostgreSQL 18.4)
  mongodb          (MongoDB 8.3.2)
  rabbitmq         (RabbitMQ 4.3.1)
  garage           (Garage S3)
  redis            (Redis 7.x)
  nginx            (Nginx 1.27.x)
```

---

## Sécurité

| Couche | Mesure |
|---|---|
| Transport | HTTPS obligatoire, HSTS, TLS 1.3 |
| Authentification | Keycloak JWT, tokens 15 min, refresh rotatif |
| Autorisation | RBAC Keycloak + Row-Level Security PostgreSQL |
| Isolation données | Un utilisateur régional ne voit que sa région via RLS |
| Validation entrées | Bean Validation côté Spring, Pydantic côté FastAPI |
| Requêtes SQL | Spring Data JPA uniquement, zéro SQL brut, zéro injection possible |
| Conformité OWASP | OWASP Top 10 + OWASP ASVS, vérification avant chaque mise en production |
| Conformité DGSSI | Guide de sécurité des applications web DGSSI Maroc |
| Audit trail | Tous les événements critiques dans MongoDB via RabbitMQ |
| Documents | Garage S3 avec accès signé temporaire (presigned URLs) |

---

## Résumé des versions

| Technologie | Version retenue |
|---|---|
| Kotlin | 2.4.0 |
| Spring Boot | 4.0.6 |
| Java (JDK) | 21 LTS |
| Python | 3.12 |
| FastAPI | 0.136.3 |
| React | 19.x |
| Vite | 8.0.16 |
| TypeScript | 5.x |
| Apache ECharts | 6.0 |
| TanStack Query | latest (juin 2026) |
| PostgreSQL | 18.4 |
| MongoDB | 8.3.2 |
| RabbitMQ | 4.3.1 |
| Keycloak | 26.6.3 |
| Garage (S3) | latest stable |
| Redis | 7.x |
| Docker | latest |
| Nginx | 1.27.x |
