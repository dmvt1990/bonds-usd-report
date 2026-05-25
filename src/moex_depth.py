"""Fetch ISS 15-min delayed depth data for bonds and enrich Bond objects."""
from concurrent.futures import ThreadPoolExecutor, as_completed
from typing import List, Optional

import requests

from .models import Bond

_ISS_URL = (
    "https://iss.moex.com/iss/engines/stock/markets/bonds/boards/TQCB"
    "/securities/{secid}.json"
    "?iss.only=marketdata&iss.meta=off"
    "&marketdata.columns=SECID,BID,OFFER,BIDDEPTHT,OFFERDEPTHT"
)
_TIMEOUT = 10
_MAX_WORKERS = 10


def _fetch_one(secid: str) -> Optional[dict]:
    try:
        r = requests.get(_ISS_URL.format(secid=secid), timeout=_TIMEOUT)
        r.raise_for_status()
        md = r.json().get("marketdata", {})
        cols = md.get("columns", [])
        rows = md.get("data", [])
        if not rows:
            return None
        return dict(zip(cols, rows[0]))
    except Exception as e:
        print(f"[depth] {secid}: {e}")
        return None


def enrich_bonds(bonds: List[Bond]) -> None:
    """Fetch ISS depth and set bid_est_usd / offer_est_usd on each bond in-place."""
    with ThreadPoolExecutor(max_workers=_MAX_WORKERS) as executor:
        futures = {executor.submit(_fetch_one, b.secid): b for b in bonds if b.secid}
        for future in as_completed(futures):
            bond = futures[future]
            row = future.result()
            if not row:
                continue
            bid = row.get("BID")
            offer = row.get("OFFER")
            biddeptht = row.get("BIDDEPTHT")
            offerdeptht = row.get("OFFERDEPTHT")
            fv = bond.facevalue
            if fv and bid and biddeptht:
                bond.bid_est_usd = biddeptht * fv * bid / 100
            if fv and offer and offerdeptht:
                bond.offer_est_usd = offerdeptht * fv * offer / 100
