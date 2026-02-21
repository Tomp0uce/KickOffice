# LiteLLM Local Proxy (Référence)

Ce dossier contient les fichiers de référence pour le proxy LiteLLM local.
Il est utilisé pour faciliter les mises à jour futures et comme documentation.

## Utilisation

Ce proxy Python sert d'intermédiaire local vers `https://litellm.kickmaker.net`.
Il ajoute automatiquement les headers d'authentification LiteLLM à chaque requête.

## Configuration

1. Copiez `.auth.env.template` en `.auth.env`
2. Remplissez vos credentials dans `.auth.env`
3. Lancez le proxy : `python proxy.py`

Le proxy écoute sur `http://localhost:4000` par défaut.

## Credentials

- `X-User-Key` : Votre clé personnelle LiteLLM (obtenue auprès de l'équipe infra)
- `X-OpenWebUi-User-Email` : Votre email Kickmaker (`prenom.nom@kickmaker.net`)

## Note

Les credentials sont maintenant gérés **directement dans l'interface KickOffice**
(onglet "Compte" dans les Paramètres). Ce dossier est conservé comme référence.
