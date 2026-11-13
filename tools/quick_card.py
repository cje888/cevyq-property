#!/usr/bin/env python3
"""Quick property cards: copy Jessica's card (sa/) for a real agent.

Only the agent details, listings and photos change. The design, layout,
booking, calculator and lead form all stay exactly as on Jessica's card.

Two templates:
  standard (blank "style") - Jessica's card as is, agent contacted on WhatsApp.
  social / cas            - same design, but the agent's own contact buttons,
                            lead form and dock send a Text (SMS) instead of
                            WhatsApp, and a row of social/site links shows
                            under the bio. Use this when the agent has no
                            WhatsApp (common in the US) or you have socials
                            to show (from a Linktree page etc). Set "style"
                            to cas or social in the Agents sheet, and put the
                            links in "socials" as Label|URL pairs, comma
                            separated, e.g. Facebook|https://fb.com/x, TikTok|https://tiktok.com/@x

  python3 tools/quick_card.py --blank agents.xlsx        make an empty sheet to fill in
  python3 tools/quick_card.py agents.xlsx photos/        build every agent in the sheet
  python3 tools/quick_card.py agents.xlsx photos/ --only erica --force

Sheet: tab "Agents" (one row per agent) and tab "Listings" (one row per home).
Photos: photos/<slug>/agent.jpg, optional hero.jpg, and the listing photos
named in the Listings tab. Output goes to <repo>/<slug>/ (index.html, agent.jpg, img/).

Nothing made up is carried over from Jessica: her reviews, sold homes, sales
figures, ticker and licence numbers are all cleared. Only facts in the sheet show.
"""
import json, re, sys, shutil, argparse
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
MASTER = REPO / "sa" / "index.html"
COUNTRY_MASTER = {"SA": "sa", "US": "usa", "AU": "aus"}
PRICE = {"SA": "R1,500", "US": "$85", "AU": "A$110"}
ROLE = {"SA": "Property Practitioner", "US": "Realtor", "AU": "Sales Agent"}
SOCIAL_STYLES = ("cas", "social")
MONEY_KEYS = ["cur", "sep", "rentPer", "rentPerWord", "loc", "tz", "tzName", "hours", "times",
              "moveIn", "moveInTxt", "calc", "types", "budgetBuy", "budgetRent", "telHint"]
CLAIM_WA = "https://wa.me/66631101888?text=Hi%20Chris%2C%20I%20saw%20the%20card%20you%20made%20for%20me.%20Let%20us%20talk."
CLAIM_MAIL = "mailto:chris@cevyq.com?subject=My%20CEVYQ%20card&body=Hi%20Chris%2C%20I%20saw%20the%20card%20you%20made%20for%20me.%20Let%20us%20talk.%20My%20name%20is%20"

AGENT_COLS = ["slug", "name", "agency", "country", "whatsapp", "email", "city", "office",
              "areas", "role", "intro", "profile link", "badge", "licence", "chips", "mark", "source",
              "style", "socials"]
LIST_COLS = ["slug", "photos", "buy or rent", "price", "beds", "baths", "garages", "size",
             "area", "type", "title", "flag", "features", "description", "link"]
AGENT_HELP = ["link name, lower case (blank = from name)", "Erica van Eden", "Pam Golding Properties",
              "SA, US or AU (blank = SA)", "with country code, e.g. 27834340515", "optional", "Hout Bay, Cape Town",
              "optional, e.g. Hout Bay office", "comma list, e.g. Hout Bay, Scott Estate", "optional (blank = country default)",
              "optional one or two lines", "agent page on agency site, optional", "optional, only if true, e.g. GoldClub Agent",
              "optional, only if true, e.g. PPRA full status FFC", "optional comma list, e.g. English, Afrikaans",
              "optional logo letters (blank = agency initials)", "optional, where photos came from, e.g. pamgolding.co.za",
              "blank = standard (Jessica, WhatsApp), or cas (Text/SMS + socials row)",
              "only for style cas, Label|URL pairs, comma separated, e.g. Facebook|https://fb.com/x"]
LIST_HELP = ["same slug as Agents", "file names in photos/<slug>/, several with +", "buy or rent", "number, blank = POA",
             "blank for land", "", "optional", "optional, e.g. 210 m² · erf 496 m²", "suburb", "optional, e.g. House",
             "optional (blank = made from beds, type, area)", "optional, e.g. New listing, Under offer",
             "optional, split with ;", "optional", "listing page, optional"]


def read_C(html):
    m = re.search(r"^var C = (\{.*\});\s*$", html, re.M)
    if not m:
        sys.exit("Could not find the card settings (var C) in the master card")
    return m, json.loads(m.group(1))


def sub_once(html, old, new):
    if html.count(old) != 1:
        sys.exit("Master card changed, could not find: " + old[:70])
    return html.replace(old, new)


def social_links(v):
    out = []
    for part in split(v):
        label, _, url = part.partition("|")
        if label.strip() and url.strip():
            out.append((label.strip(), url.strip()))
    return out


SMS_ICON = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8"><path d="M4 5h16a1 1 0 011 1v10a1 1 0 01-1 1H9l-5 4V6a1 1 0 011-1z"/><path d="M8 10h8M8 13h5"/></svg>'


def apply_social_style(html, socials):
    # Own contact buttons go by Text (SMS) instead of WhatsApp: for an agent
    # with no WhatsApp (common in the US) a wa.me link to them goes nowhere.
    # Claim-this-card button to Chris is untouched, still WhatsApp.
    html = sub_once(html,
        '<a class="act wa" href="#" data-go="wa" data-t="whatsapp"><i><svg viewBox="0 0 24 24" fill="currentColor"><path d="M12 2a10 10 0 00-8.6 15.1L2 22l5-1.3A10 10 0 1012 2zm0 18.2a8.2 8.2 0 01-4.2-1.1l-.3-.2-3 .8.8-2.9-.2-.3A8.2 8.2 0 1112 20.2zm4.5-6.1c-.2-.1-1.5-.7-1.7-.8s-.4-.1-.6.1-.7.8-.8 1-.3.2-.5.1a6.7 6.7 0 01-3.3-2.9c-.3-.4.3-.4.7-1.3.1-.2 0-.3 0-.4l-.8-1.8c-.2-.5-.4-.4-.6-.4h-.5a1 1 0 00-.7.3 3 3 0 00-.9 2.2 5.2 5.2 0 001.1 2.8 11.9 11.9 0 004.6 4c1.7.7 2.3.8 3.2.7a2.7 2.7 0 001.8-1.3 2.2 2.2 0 00.2-1.3c-.1-.1-.3-.2-.5-.3z"/></svg></i>WhatsApp</a>',
        f'<a class="act wa" href="#" data-go="sms" data-t="sms"><i>{SMS_ICON}</i>Text</a>')
    html = sub_once(html,
        'if (g === "wa") wa(HI); else if (g === "call") location.href = "tel:" + C.tel;',
        'if (g === "wa") wa(HI); else if (g === "sms") location.href = "sms:" + C.tel + "?&body=" + encodeURIComponent(HI); else if (g === "call") location.href = "tel:" + C.tel;')
    html = sub_once(html,
        '<a class="w" href="#" data-go="wa" data-t="whatsapp"><svg viewBox="0 0 24 24" fill="currentColor"><path d="M12 2a10 10 0 00-8.6 15.1L2 22l5-1.3A10 10 0 1012 2zm0 18.2a8.2 8.2 0 01-4.2-1.1l-.3-.2-3 .8.8-2.9-.2-.3A8.2 8.2 0 1112 20.2z"/></svg>Chat</a>',
        f'<a class="w" href="#" data-go="sms" data-t="sms">{SMS_ICON}Text</a>')
    html = sub_once(html, '<span id="fSendTxt">Send on WhatsApp</span>', '<span id="fSendTxt">Send by text</span>')
    html = sub_once(html, '$("fSendTxt").textContent = m === "sell" ? "Get my free valuation" : "Send on WhatsApp";',
                    '$("fSendTxt").textContent = m === "sell" ? "Get my free valuation" : "Send by text";')
    html = sub_once(html, 'wa(msg); say("Sent to " + C.first);',
                    'location.href = "sms:" + C.tel + "?&body=" + encodeURIComponent(msg); say("Sent to " + C.first);')
    if socials:
        chips = "".join(f'<a class="chip" href="{esc_attr(u)}" target="_blank" rel="noopener">{esc_attr(l)}</a>' for l, u in socials)
        html = sub_once(html, '<div class="chips" id="chips"></div>\n    <div class="actions">',
                        f'<div class="chips" id="chips"></div>\n    <div class="chips" id="socials">{chips}</div>\n    <div class="actions">')
    return html


def num(v):
    if v in (None, ""):
        return None
    s = re.sub(r"[^\d.]", "", str(v))
    if not s:
        return None
    f = float(s)
    return int(f) if f == int(f) else f


def split(v, sep=","):
    return [x.strip() for x in str(v or "").split(sep) if x.strip()]


def esc_attr(s):
    return str(s).replace("&", "&amp;").replace('"', "&quot;").replace("<", "&lt;").replace(">", "&gt;")


def slugify(s):
    return re.sub(r"[^a-z0-9]+", "-", s.lower()).strip("-")


def money_plain(p, cur, sep):
    return cur + f"{p:,}".replace(",", sep)


def resize(src, dst, width):
    from PIL import Image, ImageOps
    im = Image.open(src)
    if im.format == "JPEG" and im.width <= width:
        dst.parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(src, dst)
        return
    im = ImageOps.exif_transpose(im).convert("RGB")
    if im.width > width:
        im = im.resize((width, round(im.height * width / im.width)), Image.LANCZOS)
    dst.parent.mkdir(parents=True, exist_ok=True)
    im.save(dst, "JPEG", quality=78, optimize=True, progressive=True)


def build(a, rows, photos, force=False):
    slug = slugify(a.get("slug") or a["name"])
    out = REPO / slug
    if slug in ("sa", "usa", "aus", "malee", "erica", "dashboard", "tools"):
        sys.exit(f"{slug}: that folder is a master or existing card, pick another slug")
    if out.exists() and not force:
        sys.exit(f"{slug}: folder already exists (use --force to rebuild it)")
    pdir = Path(photos) / slug
    if not (pdir / "agent.jpg").exists():
        sys.exit(f"{slug}: missing {pdir / 'agent.jpg'}")
    if not rows:
        sys.exit(f"{slug}: no listings in the Listings tab")

    ctry = (a.get("country") or "SA").strip().upper()
    if ctry not in COUNTRY_MASTER:
        sys.exit(f"{slug}: country must be SA, US or AU")
    html = MASTER.read_text(encoding="utf-8")
    m, C = read_C(html)
    _, CM = read_C((REPO / COUNTRY_MASTER[ctry] / "index.html").read_text(encoding="utf-8"))
    for k in MONEY_KEYS:
        C[k] = CM[k]

    name = a["name"].strip(); first = name.split()[0]; agency = a["agency"].strip()
    city = (a.get("city") or "").strip()
    office = (a.get("office") or "").strip()
    areas = split(a.get("areas")) or [city]
    phone = re.sub(r"\D", "", str(a.get("whatsapp") or ""))
    if len(phone) < 9:
        sys.exit(f"{slug}: WhatsApp number needs the country code, e.g. 27834340515")
    is_social = (a.get("style") or "").strip().lower() in SOCIAL_STYLES
    socials = social_links(a.get("socials")) if is_social else []
    contact_word = "Text" if is_social else "WhatsApp"
    role = (a.get("role") or "").strip() or f"{ROLE[ctry]} · {areas[0]}"
    words = [w for w in re.findall(r"[A-Za-z]+", agency) if w.lower() not in ("the", "and", "of", "properties", "property", "realty", "real", "estate", "pty", "ltd")]
    mark = (a.get("mark") or "").strip() or "".join(w[0] for w in words[:2]).upper() or agency[:1].upper()

    # listings
    homes, n = [], 0
    for r in rows:
        n += 1
        mode = "rent" if str(r.get("buy or rent") or "").strip().lower().startswith("r") else "buy"
        imgs = split(r.get("photos"), "+")
        if not imgs:
            sys.exit(f"{slug}: listing {n} has no photo")
        beds, baths = num(r.get("beds")), num(r.get("baths"))
        area = (r.get("area") or areas[0]).strip()
        kind = (r.get("type") or "").strip()
        title = (r.get("title") or "").strip()
        if not title:
            title = (f"{beds} bed {kind.lower() or 'home'} in {area}" if beds else f"{kind or 'Property'} in {area}")
        size = str(r.get("size") or "").strip()
        feat = split(r.get("features"), ";")
        if not feat:
            feat = [x for x in [f"{beds} bedrooms" if beds else "", f"{baths} bathrooms" if baths else "",
                                f"{num(r.get('garages'))} garages" if num(r.get("garages")) else "", size, area] if x][:4]
        h = {"id": f"{slug[:3].upper()}-{n:02d}", "m": mode, "area": area, "title": title,
             "loc": f"{area}, {city.split(',')[0]}" if city and area not in city else (city or area),
             "price": num(r.get("price")) or 0, "beds": beds, "baths": baths if beds else None,
             "size": size, "flag": (r.get("flag") or "").strip() or ("To rent" if mode == "rent" else "For sale"),
             "img": [], "desc": (r.get("description") or "").strip() or f"{title}. Ask {first} for the full details or book a viewing below.",
             "feat": feat}
        if num(r.get("garages")):
            h["cars"] = num(r.get("garages"))
        if kind and not beds:
            h["kind"] = kind
        if n == 1:
            h["hot"] = 1
        if r.get("link"):
            h["src"] = str(r["link"]).strip()
        for im in imgs:
            if not (pdir / im).exists():
                sys.exit(f"{slug}: missing photo {pdir / im}")
            h["img"].append(im)
        homes.append(h)

    # rename photos to p1.jpg, p2.jpg ... so the folder stays tidy
    pmap, k = {}, 0
    for h in homes:
        for i, im in enumerate(h["img"]):
            if im not in pmap:
                k += 1; pmap[im] = f"p{k}.jpg"
            h["img"][i] = pmap[im]

    n_buy = sum(h["m"] == "buy" for h in homes); n_rent = len(homes) - n_buy
    list_areas = list(dict.fromkeys(h["area"] for h in homes))

    def plain(h):
        if not h["price"]:
            return "price on application"
        return money_plain(h["price"], C["cur"], C["sep"]) + (" " + C["rentPerWord"] if h["m"] == "rent" else "")

    C.update({
        "slug": slug, "cardid": f"property-{slug}", "preview": True, "previewDays": 7,
        "allUrl": (a.get("profile link") or "").strip(),
        "name": name, "first": first, "role": role, "where": " · ".join(areas[:3]), "mark": mark,
        "agency": agency, "city": city, "verified": (a.get("badge") or "").strip(),
        "intro": (a.get("intro") or "").strip() or f"Homes to buy{' and rent' if n_rent else ''} across {', '.join(areas[:3])}. Honest advice on price, quick replies and a clear plan from the first call to the keys.",
        "stats": [[len(homes), "listings right now", "", ""],
                  [n_rent if n_rent and n_buy else len(list_areas), "to rent" if n_rent and n_buy else ("areas" if len(list_areas) != 1 else "area"), "", ""],
                  [len(areas), "areas covered" if len(areas) != 1 else "area covered", "", ""]],
        "rating": "", "reviewsN": "", "reviews": [], "sold": [],
        "chips": split(a.get("chips")) or ((["Sales and rentals"] if n_rent else ["Property sales"]) + ([city.split(",")[0]] if city else [])),
        "licCo": agency + (", " + office if office else ""),
        "licA": (a.get("licence") or "").strip() or (city and f"Serving {city}") or "",
        "licB": (a.get("badge") or "").strip(),
        "ticker": [[h["flag"].upper(), h["title"], plain(h)] for h in homes[:5]] + [["AREAS", " · ".join(areas), ""]],
        "areas": list_areas, "homes": homes, "viewNote": "Viewings by appointment.",
        "addrHint": f"e.g. 12 Main Road, {areas[0]}",
        "consent": CM["consent"].replace(CM["agency"], agency).replace(CM["first"], first),
        "route": f"Goes straight to {first}" + (" and the office" if office else ""),
        "whyEye": f"Working with {first}", "whyH": "Why sell with me",
        "why": [["A clear plan", "Price and timing set from real recent sales in your area, not a flattering guess."],
                ["Honest answers", "Straight talk and regular updates from the first call to the final signature."],
                ["Quick replies", f"{contact_word} {first} straight from this card." if is_social else f"Message {first} on WhatsApp straight from this card."],
                ["Local knowledge", f"Working in {' and '.join(areas[:2])}, so you get advice for your street, not the whole city."]],
        "info": [["Office", agency + (", " + office if office else (", " + city if city else ""))], ["Viewings", "By appointment"]],
        "claimEye": f"Made for {first} by CEVYQ", "claimH": f"Like it, {first}?",
        "claimP": (f"This is a preview of your own digital card, built from your {agency} profile and listings. Every open and tap is counted, "
                   f"and enquiries come to you {'by text' if is_social else 'on WhatsApp'}. It can go live in a day."),
        "fine": (f"Preview made for {name} by CEVYQ Group. Listings and photos are from {(a.get('source') or 'your public listings').strip()}. "
                 f"This is not an official {agency} page. Privacy: opens and taps are counted without names. If you send an enquiry, "
                 f"your name and number go to {first} and are kept safely by CEVYQ Group so {first} can reply. To see or remove your details, "
                 "email chris@cevyq.com or WhatsApp +66 63 110 1888."),
        "wa": phone, "tel": "+" + phone, "email": (a.get("email") or "").strip(),
        "shareText": f"Homes to buy{' or rent' if n_rent else ''} in {city.split(',')[0] or areas[0]}",
        "claimWa": CLAIM_WA, "claimPrice": PRICE[ctry],
    })

    desc = f"Homes to buy{' or rent' if n_rent else ''} in {', '.join(areas[:3])}. {contact_word} {first}."
    short_city = city.split(",")[0] or areas[0]
    html = html[:m.start()] + "var C = " + json.dumps(C, ensure_ascii=False) + ";" + html[m.end():]
    html = re.sub(r"<title>.*?</title>", f"<title>{esc_attr(name)} | {esc_attr(role.split(' · ')[0])}, {esc_attr(short_city)}</title>", html, count=1)
    html = re.sub(r'(<meta name="description" content=")[^"]*', lambda x: x.group(1) + esc_attr(desc), html, count=1)
    html = re.sub(r'(<meta property="og:title" content=")[^"]*', lambda x: x.group(1) + esc_attr(f"{name} | {short_city} Property"), html, count=1)
    html = re.sub(r'(<meta property="og:description" content=")[^"]*', lambda x: x.group(1) + esc_attr(desc), html, count=1)
    html = re.sub(r'(<meta property="og:image" content=")[^"]*', lambda x: x.group(1) + f"https://property.cevyq.com/{slug}/agent.jpg", html, count=1)
    html = sub_once(html, '<meta name="robots" content="index,follow">', '<meta name="robots" content="noindex,nofollow">')
    html = sub_once(html, "from R1,500</div>", esc_attr(PRICE[ctry]) + "</div>")
    html = re.sub(r'<a class="w" href="https://wa\.me/66631101888\?text=[^"]*"', f'<a class="w" href="{CLAIM_WA}"', html, count=1)
    # no agency dashboard button on real agent cards (Chris, 8 Nov 2026)
    html, k = re.subn(r'\r?\n[ \t]*<a class="d" href="/dashboard/\?c=sa"[^\n]*?</a>', "", html)
    if k != 1:
        sys.exit("Master card changed, could not find the agency dashboard button")
    html = re.sub(r'<a class="e" href="mailto:chris@cevyq\.com\?[^"]*"', f'<a class="e" href="{CLAIM_MAIL}"', html, count=1)
    html = sub_once(html, 'say("Sample only: WhatsApp opened. No enquiry was saved to the office.");', 'say("Sent to " + C.first);')
    # Jessica's demo sends "Hello CEVYQ sales team" from every button; a real agent's card
    # messages the agent by name and keeps each button's own message
    html = sub_once(html, 'function wa(){ window.open("https://wa.me/" + C.wa + "?text=" + encodeURIComponent(HI), "_blank", "noopener"); }',
                    'function wa(text){ window.open("https://wa.me/" + C.wa + "?text=" + encodeURIComponent(text), "_blank", "noopener"); }')
    html = sub_once(html, 'var HI = "Hello CEVYQ sales team, I\'m interested in the offer.";',
                    'var HI = "Hi " + C.first + ", I found your card and would like to ask about a property.";')
    html = sub_once(html, 'encodeURIComponent("Interested in the CEVYQ offer")', 'encodeURIComponent("Property enquiry")')
    html = sub_once(html, "<!-- Sample card: production tracking and lead ingestion are disabled. -->", "")
    html, k = re.subn(r'<script src="https://cevyq-stats\.chris-e41\.workers\.dev/t\.js"[^>]*></script>',
                      lambda _: f'<script src="https://cevyq-stats.chris-e41.workers.dev/t.js" data-card="property-{slug}" '
                                f'data-name="{esc_attr(name)}, {esc_attr(agency)} (preview)" data-agency="" data-person="{esc_attr(name)}" data-agent="Chris" defer></script>', html)
    if k != 1:
        sys.exit("Master card changed, could not find the Card Desk tag")
    if is_social:
        html = apply_social_style(html, socials)

    if out.exists():
        shutil.rmtree(out)
    (out / "img").mkdir(parents=True)
    (out / "index.html").write_text(html, encoding="utf-8", newline="")
    resize(pdir / "agent.jpg", out / "agent.jpg", 700)
    hero = pdir / "hero.jpg"
    resize(hero if hero.exists() else pdir / next(iter(pmap)), out / "img" / "hero.jpg", 1600)
    for src, dst in pmap.items():
        resize(pdir / src, out / "img" / dst, 1200)
    # Dashboard stats page: same generic files as Jessica's, no per-card edits needed.
    shutil.copytree(MASTER.parent / "stats", out / "stats")
    return slug


def read_sheet(path):
    import openpyxl
    wb = openpyxl.load_workbook(path, data_only=True)

    def rows(tab):
        ws = wb[tab]
        it = ws.iter_rows(values_only=True)
        head = [str(h or "").strip().lower() for h in next(it)]
        out = []
        for r in it:
            d = {head[i]: ("" if v is None else str(v).strip() if not isinstance(v, (int, float)) else v) for i, v in enumerate(r) if i < len(head) and head[i]}
            if any(str(v).strip() for v in d.values()) and not str(d.get("slug", "")).startswith("("):
                out.append(d)
        return out
    return rows("Master" if "Master" in wb.sheetnames else "Agents"), rows("Listings")


def blank(path):
    import openpyxl
    from openpyxl.styles import Font, PatternFill
    wb = openpyxl.Workbook()
    for i, (tab, cols, helps) in enumerate([("Agents", AGENT_COLS, AGENT_HELP), ("Listings", LIST_COLS, LIST_HELP)]):
        ws = wb.active if i == 0 else wb.create_sheet()
        ws.title = tab
        ws.append(cols)
        ws.append(["(" + h + ")" if h else "" for h in helps])
        for c in ws[1]:
            c.font = Font(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor="0D3B46")
        for c in ws[2]:
            c.font = Font(italic=True, color="888888")
        for col in ws.columns:
            ws.column_dimensions[col[0].column_letter].width = 22
        ws.freeze_panes = "A2"
    wb.save(path)
    print("Blank sheet saved:", path)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("sheet")
    ap.add_argument("photos", nargs="?")
    ap.add_argument("--blank", action="store_true")
    ap.add_argument("--only", nargs="*")
    ap.add_argument("--force", action="store_true")
    o = ap.parse_args()
    if o.blank:
        return blank(o.sheet)
    if not o.photos:
        sys.exit("Give the photos folder too")
    agents, listings = read_sheet(o.sheet)
    for a in agents:
        slug = slugify(a.get("slug") or a.get("name", ""))
        if o.only and slug not in o.only:
            continue
        rows = [r for r in listings if slugify(str(r.get("slug", ""))) == slug]
        print("Built", build(a, rows, o.photos, o.force), "with", len(rows), "listings ->", f"https://property.cevyq.com/{slug}/")


if __name__ == "__main__":
    main()
