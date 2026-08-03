"""
e-Boekhouden SOAP koppeling
---------------------------
Kleine wrapper rond de (oudere, goed gedocumenteerde) SOAP-API van e-Boekhouden.nl,
gebruikt om verwerkte bonnetjes als boeking (mutatie) weg te schrijven.
 
Belangrijk: dit schrijft direct een mutatie weg zodra je 'm aanroept - er bestaat geen
apart "concept"-mutatie-type in de SOAP-API van e-Boekhouden. Daarom laten we de
controle in onze eigen app plaatsvinden (het overzicht dat je bekijkt en corrigeert
vóórdat je op "Boek in e-Boekhouden" klikt) - pas na die controle wordt deze functie
aangeroepen.
 
Elke klant-administratie heeft eigen inloggegevens (Gebruikersnaam, SecurityCode1,
SecurityCode2), die per klant als aparte environment variables in Railway staan.
"""
 
import os
from datetime import datetime
from zeep import Client
 
WSDL_URL = "https://soap.e-boekhouden.nl/soap.asmx?wsdl"
 
# BTW-percentage -> BTW-code voor inkopen (zie e-Boekhouden SOAP documentatie)
BTW_CODE_INKOOP = {
    21: "HOOG_INK_21",
    9: "LAAG_INK",
    0: "GEEN",
}
 
 
class EBoekhoudenError(Exception):
    pass
 
 
def _get_credentials(klant_prefix):
    """
    Haalt gebruikersnaam + beide securitycodes op voor een klant, via environment
    variables met een klant-specifiek voorvoegsel, bijvoorbeeld:
    EBOEKHOUDEN_BURGERME_USERNAME, EBOEKHOUDEN_BURGERME_SEC1, EBOEKHOUDEN_BURGERME_SEC2
    """
    username = os.environ.get(f"EBOEKHOUDEN_{klant_prefix}_USERNAME")
    sec1 = os.environ.get(f"EBOEKHOUDEN_{klant_prefix}_SEC1")
    sec2 = os.environ.get(f"EBOEKHOUDEN_{klant_prefix}_SEC2")
    if not (username and sec1 and sec2):
        raise EBoekhoudenError(
            f"Inloggegevens voor '{klant_prefix}' ontbreken (environment variables "
            f"EBOEKHOUDEN_{klant_prefix}_USERNAME / _SEC1 / _SEC2 niet allemaal ingesteld)"
        )
    return username, sec1, sec2
 
 
def _zoek_relatiecode_op_naam(soap_client, session_id, sec2, naam):
    """
    Zoekt een relatie op in e-Boekhouden via de bedrijfsnaam (Trefwoord), en geeft
    de bijbehorende relatiecode terug als er precies één match is. Geeft None
    terug als er geen (of meerdere onduidelijke) matches zijn.
    """
    if not naam:
        return None
    try:
        result = soap_client.service.GetRelaties(session_id, sec2, {"Trefwoord": naam})
    except Exception:
        return None
    if getattr(result, "LastErrorCode", None):
        return None
    relaties_container = getattr(result, "Relaties", None)
    if not relaties_container:
        return None
    relaties = getattr(relaties_container, "cRelatie", None)
    if not relaties:
        return None
    if not isinstance(relaties, list):
        relaties = [relaties]
    if len(relaties) == 1:
        return relaties[0].Code
    return None  # meerdere matches: niet gokken, gebruiker moet zelf kiezen
 
 
def boek_bonnetje(klant_prefix, leverancier, factuurdatum_ddmmjjjj, factuurnummer,
                   omschrijving, betaalrekening, regels, relatiecode="", betalingstermijn="0",
                   crediteurenrekening="1700", eu_leverancier=False, kruispost=None,
                   zoeknaam_relatie=None):
    """
    Boekt één factuur of bonnetje (met mogelijk meerdere BTW-regels).
 
    - Heeft de factuur een factuurnummer? -> Soort "FactuurOntvangen", geboekt op de
      crediteurenrekening (standaard 1700). Vereist een relatie (leverancier) die al
      bestaat in e-Boekhouden. Is er geen relatiecode meegegeven, dan zoekt de app
      'm automatisch op via de leveranciersnaam (GetRelaties) - je hoeft dus niet
      voor elke leverancier handmatig een nummer te verzinnen.
    - Geen factuurnummer (kassabonnetje)? -> Soort "GeldUitgegeven", direct van de
      gekozen betaalrekening (kas/bank/pin).
 
    regels: lijst van dicts met keys:
        - btw_percentage (21, 9 of 0)
        - bedrag_excl_btw
        - btw_bedrag
        - kostenrekening (grootboekcode, bijv. "7012")
 
    eu_leverancier: True voor buitenlandse (EU) leveranciers - gebruikt de
    "Leveringen/diensten van binnen EU 0%" BTW-code (verlegde BTW) i.p.v. de
    gewone "Geen btw"-code.
 
    kruispost: optioneel dict {"kostenrekening": ..., "kruispost_rekening": ...}
    voor bezorgplatforms (Thuisbezorgd/Uber Eats/Mollie) - naast de gewone
    kostenregel(s) wordt een extra, negatieve verrekeningsregel toegevoegd tegen
    de kruispost-rekening, gelijk aan het totale bedrag incl. BTW.
 
    Geeft het/de mutatienummer(s) terug die e-Boekhouden aanmaakt.
    """
    heeft_factuurnummer = bool((factuurnummer or "").strip())
    soort = "FactuurOntvangen" if heeft_factuurnummer else "GeldUitgegeven"
 
    username, sec1, sec2 = _get_credentials(klant_prefix)
 
    soap_client = Client(WSDL_URL)
    session_id = None
    mutatienummers = []
 
    try:
        open_result = soap_client.service.OpenSession(username, sec1, sec2)
        if getattr(open_result, "LastErrorCode", None):
            raise EBoekhoudenError(f"Kon geen sessie openen: {open_result.LastErrorDescription}")
        session_id = open_result.SessionID
 
        # Geen relatiecode meegegeven? Probeer 'm automatisch op te zoeken op naam.
        if soort == "FactuurOntvangen" and not (relatiecode or "").strip():
            zoeknaam = zoeknaam_relatie or leverancier
            relatiecode = _zoek_relatiecode_op_naam(soap_client, session_id, sec2, zoeknaam) or ""
 
        if soort == "FactuurOntvangen" and not relatiecode:
            raise EBoekhoudenError(
                f"Kon geen (eenduidige) relatie vinden voor '{leverancier}' in e-Boekhouden. "
                "Maak deze leverancier eerst aan als relatie in e-Boekhouden (Relaties), "
                "of vul de relatiecode handmatig in bij het controleren."
            )
 
        # Datum van DD-MM-JJJJ naar JJJJ-MM-DD (wat e-Boekhouden verwacht)
        try:
            dag, maand, jaar = factuurdatum_ddmmjjjj.split("-")
            datum_iso = f"{jaar}-{maand}-{dag}"
        except Exception:
            datum_iso = datetime.now().strftime("%Y-%m-%d")
 
        mutatie_regels = []
        totaal_incl_btw = 0.0
        for regel in regels:
            btw_pct = int(regel.get("btw_percentage") or 0)
            if eu_leverancier:
                btw_code = "BI_EU_INK"  # Leveringen/diensten van binnen EU (verlegd)
            else:
                btw_code = BTW_CODE_INKOOP.get(btw_pct, "GEEN")
            excl = float(regel.get("bedrag_excl_btw") or 0)
            btw_bedrag = float(regel.get("btw_bedrag") or 0)
            incl = excl + btw_bedrag
            totaal_incl_btw += incl
            mutatie_regels.append({
                "BedragInvoer": incl,
                "BedragExclBTW": excl,
                "BedragBTW": btw_bedrag,
                "BedragInclBTW": incl,
                "BTWCode": btw_code,
                "BTWPercentage": btw_pct,
                "TegenrekeningCode": regel.get("kostenrekening"),
            })
 
        # Bezorgplatform-patroon: extra negatieve verrekeningsregel tegen de
        # kruispost-rekening, gelijk aan het totaal incl. BTW van de kostenregel(s).
        if kruispost:
            mutatie_regels.append({
                "BedragInvoer": -totaal_incl_btw,
                "BedragExclBTW": -totaal_incl_btw,
                "BedragBTW": 0,
                "BedragInclBTW": -totaal_incl_btw,
                "BTWCode": "GEEN",
                "BTWPercentage": 0,
                "TegenrekeningCode": kruispost.get("kruispost_rekening"),
            })
 
        rekening = crediteurenrekening if soort == "FactuurOntvangen" else betaalrekening
 
        oMut = {
            "Soort": soort,
            "Datum": datum_iso,
            "Rekening": rekening,
            "RelatieCode": relatiecode or "",
            "Factuurnummer": factuurnummer or "",
            "Omschrijving": f"{leverancier or ''} - {omschrijving or ''}".strip(" -"),
            "Betalingstermijn": betalingstermijn or "0",
            "InExBTW": "EX",
            "MutatieRegels": {"cMutatieRegel": mutatie_regels},
        }
 
        add_result = soap_client.service.AddMutatie(session_id, sec2, oMut)
        if getattr(add_result, "LastErrorCode", None):
            raise EBoekhoudenError(f"Boeken mislukt: {add_result.LastErrorDescription}")
        mutatienummers.append(add_result.Mutatienummer)
 
        return mutatienummers
 
    finally:
        if session_id:
            try:
                soap_client.service.CloseSession(session_id)
            except Exception:
                pass
 
 
