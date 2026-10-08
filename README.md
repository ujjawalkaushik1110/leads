# Lead Scraper v2 + local AI (Ollama) + web interface - no paid services

## Easiest way: the web interface
    pip install -r requirements.txt
    python lead_scraper.py ui          # opens http://127.0.0.1:8501 (keep lead_scraper.py and lead_ui.py in the same folder)

1. **Setup** - connect Ollama (it can start it for you), pick the model from a dropdown (embedding models hidden), see a RAM-based model recommendation,
   download a model with a progress bar, test it, choose AI mode (auto / always / off). Connect Google and Reddit logins if you want those sources.
2. **Run** - choose *Find leads*, *Crawl a website* or *Reddit signals*, fill the form (every field has a hover tooltip), press Run, watch live progress and logs,
   press *Stop* to export what was found so far.
3. **Results** - sortable/filterable table, click a row for every field (AI reason, people, outreach draft with a Copy button), download Excel / CSV / JSONL, re-download earlier runs.

The forms are generated from the command-line options, so the UI and the CLI always offer exactly the same features.
Security: the UI listens on 127.0.0.1 only, validates Host/Origin headers and requires a random per-session token. Scraped text is always shown as inert text, never as HTML.

## Command line (same engine)

    pip install -r requirements.txt

## Mode 1: find leads for an industry
    python lead_scraper.py leads "dental clinics" --location Mumbai --depth 3 --max-leads 300
    python lead_scraper.py leads "SaaS startups" "digital agencies" --engines bing,brave,google --no-osm
    python lead_scraper.py leads "solar installers" --include "rooftop,installation" --exclude "jobs" --tlds in,com --require-contact

## Mode 2: crawl a whole website
    python lead_scraper.py crawl https://example.com --max-pages 500 --depth 6 --export-text
    python lead_scraper.py crawl urls.txt --js

## AI with Ollama (automatic)
Install Ollama (https://ollama.com) once. The tool then connects by itself: it starts `ollama serve` if needed,
skips embedding models and picks the best installed chat model (qwen2.5 > llama3.x > gemma3 > mistral ...).
If Ollama is missing it just runs without AI. Everything stays on your machine.

    ollama pull qwen2.5:7b                       # or: add --pull qwen2.5:7b to any command
    python lead_scraper.py leads --ask "small dental clinics in Pune that offer implants" --outreach "website redesign"
    python lead_scraper.py leads "SaaS startups" --icp "B2B SaaS, 10-50 staff, hiring sales" --min-fit 6
    python lead_scraper.py crawl https://example.com --extract "pricing plans, team members, refund policy"

| AI feature | What it does |
|---|---|
| `--ask "..."` | Turns plain English into search queries, keywords, country TLDs, wrong-result terms and an ideal-customer profile |
| AI Fit / Reason | Reads each company's pages and scores 0-10 against your profile, with a written reason (replaces keyword relevance in Lead Score) |
| Not-a-company check | Flags directories/blogs/listicles that slipped through |
| Industry, size, summary, people, buying signals | Extracted only from site text, never invented |
| `--outreach "what you sell"` | Personalised cold-email draft per lead that has an email (review before sending) |
| `crawl` AI Insights sheet | Summary + fields of your choice (`--extract`) from the site's key pages |
| Controls | `--ai off/on/auto`, `--model`, `--ollama-host`, `--ai-top N` (limit leads judged), `--ai-workers`, `--min-fit` |

Page text is treated as untrusted data in every prompt (prompt-injection guard). Small local models can be wrong: use AI Reason to audit.

## Google Maps + Reddit (official APIs, YOUR login)
These use the platforms' own APIs with your account. Passwords never touch this tool: you sign in on Google's / Reddit's own page in your browser.

    # Google: create an OAuth client of type "Desktop app" (console.cloud.google.com > APIs & Services > Credentials),
    # enable "Places API (New)" and billing on that project. Then:
    python lead_scraper.py auth google --client-file client_secret.json --project YOUR_PROJECT_ID
    #   (simpler alternative: python lead_scraper.py auth google --api-key YOUR_KEY)

    # Reddit: create an app at reddit.com/prefs/apps, redirect URI http://127.0.0.1:8765/callback. Then:
    python lead_scraper.py auth reddit --client-id ID --secret SECRET --username YOUR_REDDIT_NAME
    #   (Reddit's sign-in page offers "Continue with Google")

    python lead_scraper.py auth status          # what is connected
    python lead_scraper.py auth logout all      # delete saved logins

    python lead_scraper.py leads "dental clinics" --location Mumbai --maps            # web + OSM + Google Maps together
    python lead_scraper.py leads "dental clinics" --location Mumbai --maps --no-web   # Google Maps only
    python lead_scraper.py reddit "looking for a dentist" "recommend implant clinic" --subs india,mumbai --days 30 --offer "dental marketing"

- Google Maps adds rating, review count, Maps link, categories and phone/website (even for businesses with no website). Google returns at most ~60 places per query, so use several queries/areas for more.
- Reddit returns posts where people ask for recommendations/help (buying-intent signals) with intent score, and with Ollama: relevance, urgency and a suggested public reply angle. It collects public usernames only: no profile scraping, no DMs.
- Logins are stored in `~/.lead_scraper/credentials.json` (owner-only permissions).

**Read before you export:** Google's Maps Platform terms restrict storing Places content beyond place IDs, so exporting Maps data into a lead spreadsheet may breach your agreement with Google. Reddit's Data API terms require approval for commercial use and bar redistribution. These are your account's contracts: check them first. LinkedIn, Facebook and Instagram have no official lead-search API, so they are not included.

## Features (also listed on the Guide sheet inside every Excel file; hover any column header for its tooltip)
| Feature | What it does |
|---|---|
| Multi-engine parallel search | DuckDuckGo, Bing, Brave, Google, Yahoo, Mojeek, Startpage, Yandex queried at the same time, merged with Reciprocal Rank Fusion. Failing engines auto-disabled. |
| Research depth 1/2/3 | 1 = your query; 2 = + contact/phone/top/services variants; 3 = + operator queries (inurl:contact, intitle:"contact us"), quote/pricing/founder/careers variants |
| OpenStreetMap | Free local-business source with phone/email/address, even for businesses with no website |
| List harvesting | Opens "Top 10..." and directory pages, harvests the outbound company links |
| Per-lead crawl | Homepage + contact/about/team pages: emails (incl. [at]/[dot]), validated phones, JSON-LD address, socials, tech stack, contact form |
| Reasoning filters | Page-type check (company vs directory/blog), --include/--exclude terms, --tlds, --exclude-domains, --verify-mx |
| Explainable score | Lead Score + Score Breakdown columns |
| Full-site crawler | Sitemap + BFS links; Pages, Contacts, Files, External Links sheets; issues audit (broken, thin, duplicate, noindex, missing title/H1/meta); --export-text |
| Resume | .jsonl checkpoint; rerun with --resume |
| Optional JS rendering | --js uses Playwright for JavaScript-only sites |

## Politeness (not configurable)
robots.txt and Crawl-delay are obeyed, requests are rate-limited per host, 429/5xx back off.
Bot-check/CAPTCHA pages are reported as "bot-protected" and skipped, not evaded.
Use only public business contact data; follow GDPR / India DPDP / anti-spam law.
