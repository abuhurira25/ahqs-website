# Search discovery maintenance

Production origin: https://ahqshealthcare.com

## Before committing page changes

Run `python3 .github/scripts/search_discovery.py sitemap` and then
`python3 .github/scripts/search_discovery.py validate`.
Commit the generated sitemap together with the page changes.

The generator keeps self-canonical, indexable HTML pages only, includes nested
pages when appropriate, and uses actual Git modification dates. Thank-you,
offline and error pages are excluded. `/readiness-hub/` is an alternate copy of
`/resources.html` and declares the latter as canonical.

## After publication

The `Search engine discovery` GitHub Actions workflow runs after a successful
GitHub Pages build. It verifies the live site-hosted IndexNow proof before
notifying the global IndexNow endpoint. Initial proof publication sends all
canonical URLs. Later builds send the latest commit's changed pages; shared
asset changes notify all canonical pages.

For a multi-commit deployment, major update, or manual recovery, run the
workflow with `submit_all=true` after confirming the deployment is live.
The workflow has read-only repository permissions and does not modify content.

The proof in `ahqs-indexnow.txt` is intentionally website-accessible for domain
verification. It is not a bank, email or search-console account credential.

HTTP 200 means IndexNow received the submission. HTTP 202 means the submission
was received but ownership validation is pending. Neither confirms indexing or
ranking. No repeated scheduled notifications are configured.

Google does not use this workflow. Its verified Search Console domain property
is `sc-domain:ahqshealthcare.com`, with the submitted sitemap at
https://ahqshealthcare.com/sitemap.xml. Use sitemap resubmission when appropriate;
do not use Google's restricted Indexing API for ordinary consultancy pages.

Reference: https://www.indexnow.org/documentation
