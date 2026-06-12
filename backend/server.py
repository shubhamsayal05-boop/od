"""ODRIV backend — FastAPI port of the Excel/VBA tool ODRIV v29.2.1 AT."""
import json
import os
import io
import logging
import tempfile
import uuid
from datetime import datetime, timezone
from pathlib import Path
from typing import Optional

from dotenv import load_dotenv
from fastapi import FastAPI, APIRouter, HTTPException, UploadFile, File, Body
from fastapi.responses import FileResponse
from motor.motor_asyncio import AsyncIOMotorClient
from starlette.middleware.cors import CORSMiddleware

from engine import config_loader, scoring, importer
from engine.classifier import classify_event, get_channel
from engine import reports as report_builder

ROOT_DIR = Path(__file__).parent
load_dotenv(ROOT_DIR / '.env')

mongo_url = os.environ['MONGO_URL']
client = AsyncIOMotorClient(mongo_url)
db = client[os.environ['DB_NAME']]

app = FastAPI(title="ODRIV")
api = APIRouter(prefix="/api")

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("odriv")

ODRIV_VERSION = "VERSION v29_2_1_AT (Python)"

PROJECT_FIELDS = ["name_code", "mode", "fuel", "gears", "software_milestone",
                  "priority", "version", "odriv_milestone", "area",
                  "target_vehicle", "number_of_gears"]


# --------------------------------------------------------------- helpers

async def moniteur(msg):
    """Macro_Log equivalent."""
    await db.macro_log.insert_one({
        "id": str(uuid.uuid4()),
        "ts": datetime.now(timezone.utc).isoformat(),
        "message": msg})


async def get_config(section=None):
    if section:
        doc = await db.config.find_one({"section": section}, {"_id": 0})
        if not doc:
            raise HTTPException(404, f"Unknown config section {section}")
        return doc["data"]
    cfg = {}
    async for doc in db.config.find({}, {"_id": 0}):
        cfg[doc["section"]] = doc["data"]
    return cfg


async def get_project():
    return await db.project.find_one({}, {"_id": 0})


def canon_map_of(cfg):
    return config_loader.build_canon_map(cfg["structure"], cfg["catalog"])


async def seed_if_needed(force=False):
    count = await db.config.count_documents({})
    if count > 0 and not force:
        return
    await db.config.delete_many({})
    sections = config_loader.load_seed_sections()
    for name, data in sections.items():
        await db.config.insert_one({"section": name, "data": data})
    await moniteur("Configuration seeded from ODRIV_v29_2_1_AT workbook")
    logger.info("Seeded %d config sections", len(sections))


@app.on_event("startup")
async def startup():
    await seed_if_needed()


# --------------------------------------------------------------- state

@api.get("/state")
async def state():
    project = await get_project()
    cfg_catalog = await get_config("catalog")
    pipeline = [{"$match": {"sdv": {"$ne": ""}}},
                {"$group": {"_id": "$sdv", "n": {"$sum": 1}}}]
    counts = {d["_id"]: d["n"] async for d in db.events.aggregate(pipeline)}
    order = {c["name"]: c["order"] for c in cfg_catalog}
    sheets = [{"name": k, "events": v} for k, v in counts.items() if v > 2]
    sheets.sort(key=lambda s: order.get(s["name"], 999))
    total_events = await db.events.count_documents({})
    rating = await db.rating_global.find_one({}, {"_id": 0})
    logs = await db.macro_log.find({}, {"_id": 0}).sort("ts", -1).limit(5).to_list(5)
    return {"project": project, "sheets": sheets, "total_events": total_events,
            "has_rating": rating is not None, "version": ODRIV_VERSION,
            "log_tail": logs}


# --------------------------------------------------------------- project

@api.post("/project/new")
async def new_project(payload: dict = Body(...)):
    await _erase_all()
    project = {f: payload.get(f) for f in PROJECT_FIELDS}
    project["id"] = str(uuid.uuid4())
    project["created_at"] = datetime.now(timezone.utc).isoformat()
    await db.project.insert_one(dict(project))
    await moniteur(f"New project has been created : {project.get('name_code')}")
    project.pop("_id", None)
    return project


@api.put("/project")
async def update_project(payload: dict = Body(...)):
    project = await get_project()
    if not project:
        raise HTTPException(400, "No project. Use NEW PROJECT first.")
    updates = {k: v for k, v in payload.items() if k in PROJECT_FIELDS}
    await db.project.update_one({"id": project["id"]}, {"$set": updates})
    if "version" in updates or "fuel" in updates:
        await moniteur("Targets have been updated (drive version / fuel changed)")
    return await get_project()


async def _erase_all():
    await db.events.delete_many({})
    await db.sdv_results.delete_many({})
    await db.rating_global.delete_many({})
    await db.project.delete_many({})


@api.delete("/project")
async def erase_all():
    await _erase_all()
    await moniteur("ERASE ALL DATA : project deleted")
    return {"ok": True}


# --------------------------------------------------------------- import

async def _store_events(parsed, filename, cfg):
    canon = canon_map_of(cfg)
    rules = cfg["definitions"]
    docs = []
    per_sdv = {}
    unclassified = 0
    for channels in parsed:
        sdv = classify_event(channels, rules)
        name = config_loader.canon_name(sdv, canon) if sdv else ""
        if not name:
            unclassified += 1
        else:
            per_sdv[name] = per_sdv.get(name, 0) + 1
        docs.append({"id": str(uuid.uuid4()), "sdv": name, "channels": channels,
                     "file": filename,
                     "imported_at": datetime.now(timezone.utc).isoformat()})
    if docs:
        await db.events.insert_many(docs)
    return {"imported": len(docs), "classified": len(docs) - unclassified,
            "unclassified": unclassified, "per_sdv": per_sdv}


@api.post("/import/file")
async def import_file(file: UploadFile = File(...)):
    project = await get_project()
    if not project:
        raise HTTPException(400, "Project information missing. Create a project first.")
    missing = [f for f in ("fuel", "gears", "software_milestone", "odriv_milestone")
               if not project.get(f)]
    if missing:
        raise HTTPException(400, "Project information missing : " + ", ".join(missing))
    content = await file.read()
    try:
        parsed, _ = importer.parse_trie(content)
    except ValueError as e:
        raise HTTPException(400, str(e))
    cfg = await get_config()
    result = await _store_events(parsed, file.filename, cfg)
    await moniteur(f"File {file.filename} has been added to your project "
                   f"({result['classified']}/{result['imported']} events classified)")
    return result


@api.post("/import/demo")
async def import_demo():
    project = await get_project()
    if not project:
        raise HTTPException(400, "Project information missing. Create a project first.")
    cfg = await get_config()
    canon = canon_map_of(cfg)
    events = importer.generate_sample_events(cfg["definitions"], cfg["structure"], canon)
    result = await _store_events(events, "DEMO_ACQUISITION.xlsx", cfg)
    await moniteur(f"Demo acquisition generated and added "
                   f"({result['classified']}/{result['imported']} events classified)")
    return result


@api.get("/import/sample")
async def download_sample():
    cfg = await get_config()
    canon = canon_map_of(cfg)
    events = importer.generate_sample_events(cfg["definitions"], cfg["structure"], canon)
    path = os.path.join(tempfile.gettempdir(), "ODRIV_sample_acquisition.xlsx")
    importer.write_sample_workbook(path, events)
    return FileResponse(path, filename="ODRIV_sample_acquisition.xlsx",
                        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# --------------------------------------------------------------- rating

@api.post("/rating/calculate")
async def calculate_rating():
    project = await get_project()
    if not project:
        raise HTTPException(400, "No project")
    cfg = await get_config()
    events = await db.events.find({"sdv": {"$ne": ""}}, {"_id": 0}).to_list(100000)
    if not events:
        raise HTTPException(400, "No events in database. Add a file first.")
    canon = canon_map_of(cfg)
    targets_lookup = config_loader.build_targets_lookup(cfg["targets"], project)
    updates, sdv_results, global_results = scoring.calculate_rating(
        project, events, cfg, targets_lookup, canon)
    # persist
    for ev_id, parts in updates.items():
        await db.events.update_one({"id": ev_id}, {"$set": parts})
    await db.sdv_results.delete_many({})
    if sdv_results:
        await db.sdv_results.insert_many([dict(r) for r in sdv_results])
    global_results["calculated_at"] = datetime.now(timezone.utc).isoformat()
    await db.rating_global.delete_many({})
    await db.rating_global.insert_one(dict(global_results))
    await moniteur("Rating has been calculated.")
    global_results.pop("_id", None)
    return {"global": global_results, "sdv_count": len(sdv_results)}


@api.get("/rating")
async def get_rating():
    project = await get_project()
    glob = await db.rating_global.find_one({}, {"_id": 0})
    rows = await db.sdv_results.find({}, {"_id": 0}).sort("order", 1).to_list(200)
    catalog = await get_config("catalog")
    return {"project": project, "global": glob, "rows": rows, "catalog": catalog,
            "version": ODRIV_VERSION}


@api.post("/rating/set-as-target")
async def set_as_target():
    project = await get_project()
    if not project:
        raise HTTPException(400, "No project")
    rows = await db.sdv_results.find({}, {"_id": 0}).to_list(200)
    if not rows:
        raise HTTPException(400, "Calculate the rating first")
    tv = await get_config("target_vehicle")
    vehicle = project.get("name_code") or "TESTED VEHICLE"
    version = "V" + str(project.get("version") or "4.6").lstrip("Vv")
    by_sdv = {r["sdv"].strip().upper(): r for r in tv["rows"]}
    for r in rows:
        d = (r.get("driv") or {}).get("index")
        dy = (r.get("dyn") or {}).get("index")
        rec = by_sdv.get(r["name"].strip().upper())
        if rec:
            rec.update({"driv": d, "dyn": dy, "vehicle": vehicle,
                        "drive_version": version})
        else:
            tv["rows"].append({"sdv": r["name"], "drive_version": version,
                               "vehicle": vehicle, "mode": project.get("mode") or "AUTO",
                               "driv": d, "dyn": dy})
    await db.config.update_one({"section": "target_vehicle"}, {"$set": {"data": tv}})
    await moniteur(f"SET AS TARGET : current results promoted to TARGET VEHICLE ({vehicle})")
    return {"ok": True, "vehicle": vehicle, "rows": len(rows)}


# --------------------------------------------------------------- SDV sheet detail

@api.get("/sdv/{name}")
async def sdv_detail(name: str):
    project = await get_project()
    cfg = await get_config()
    canon = canon_map_of(cfg)
    cname = config_loader.canon_name(name, canon)
    events = await db.events.find({"sdv": cname}, {"_id": 0}).to_list(10000)
    if not events:
        raise HTTPException(404, f"No events for SDV {name}")
    result = await db.sdv_results.find_one({"name": cname}, {"_id": 0})
    spec = cfg["structure"].get(cname, {"columns": [], "data": [], "criteria": []})
    targets_lookup = config_loader.build_targets_lookup(cfg["targets"], project or {})
    targets = targets_lookup.get(cname.strip().upper(), {})
    charts = {k.strip().upper(): v for k, v in cfg["chart_params"].items()}.get(
        cname.strip().upper(), [])
    # flatten event rows for the sheet grid
    rows = []
    for ev in events:
        ch = ev["channels"]
        row = {"id": ev["id"], "file": ev.get("file"),
               "driv": ev.get("driv"), "dyn": ev.get("dyn"), "channels": ch}
        rows.append(row)
    return {"name": cname, "result": result, "structure": spec,
            "targets": {k: {"wl": v.get("wl"), "t": v.get("t"),
                            "driv": v.get("driv"), "resp": v.get("resp")}
                        for k, v in targets.items()},
            "charts": charts, "events": rows, "project": project}


# --------------------------------------------------------------- events DB (OPEN DATABASE)

@api.get("/events")
async def list_events(search: Optional[str] = None, sdv: Optional[str] = None,
                      skip: int = 0, limit: int = 50):
    q = {}
    if sdv:
        q["sdv"] = sdv
    total = await db.events.count_documents(q)
    cur = db.events.find(q, {"_id": 0}).skip(skip).limit(min(limit, 500))
    events = await cur.to_list(min(limit, 500))
    if search:
        s = search.lower()
        events = [e for e in events if s in json.dumps(e["channels"]).lower()
                  or s in e["sdv"].lower()]
    return {"total": total, "events": events}


@api.put("/events/{event_id}")
async def update_event(event_id: str, payload: dict = Body(...)):
    ev = await db.events.find_one({"id": event_id}, {"_id": 0})
    if not ev:
        raise HTTPException(404, "Event not found")
    channels = payload.get("channels")
    if channels:
        clean = {k.replace(".", "\u00b7"): v for k, v in channels.items()}
        await db.events.update_one({"id": event_id}, {"$set": {"channels": clean}})
        await moniteur(f"Event {event_id[:8]} updated in database")
    return await db.events.find_one({"id": event_id}, {"_id": 0})


@api.delete("/events/{event_id}")
async def delete_event(event_id: str):
    res = await db.events.delete_one({"id": event_id})
    if res.deleted_count == 0:
        raise HTTPException(404, "Event not found")
    await moniteur(f"Event {event_id[:8]} deleted from database")
    return {"ok": True}


# --------------------------------------------------------------- config sheets

@api.get("/config")
async def all_config():
    return await get_config()


@api.get("/config/{section}")
async def one_config(section: str):
    return await get_config(section)


@api.put("/config/{section}")
async def put_config(section: str, payload: dict = Body(...)):
    doc = await db.config.find_one({"section": section})
    if not doc:
        raise HTTPException(404, f"Unknown config section {section}")
    await db.config.update_one({"section": section},
                               {"$set": {"data": payload.get("data")}})
    await moniteur(f"Configuration sheet '{section}' modified")
    return {"ok": True}


@api.post("/config/reset")
async def reset_config():
    await seed_if_needed(force=True)
    return {"ok": True}


@api.get("/layout/{name}")
async def layout(name: str):
    try:
        return config_loader.load_layout(name)
    except KeyError:
        raise HTTPException(404, "Unknown layout")


# --------------------------------------------------------------- logs

@api.get("/logs")
async def logs(limit: int = 200):
    rows = await db.macro_log.find({}, {"_id": 0}).sort("ts", -1).limit(limit).to_list(limit)
    return rows


# --------------------------------------------------------------- reports

@api.post("/report/{fmt}")
async def create_report(fmt: str, payload: dict = Body(default={})):
    if fmt not in ("pptx", "pdf"):
        raise HTTPException(400, "fmt must be pptx or pdf")
    project = await get_project()
    glob = await db.rating_global.find_one({}, {"_id": 0})
    rows = await db.sdv_results.find({}, {"_id": 0}).sort("order", 1).to_list(200)
    if not project or not rows:
        raise HTTPException(400, "Calculate the rating before creating a report")
    cfg = await get_config()
    charts_cfg = {k.strip().upper(): v for k, v in cfg["chart_params"].items()}
    charts = {}
    for r in rows:
        events = await db.events.find({"sdv": r["name"]}, {"_id": 0}).to_list(5000)
        params = charts_cfg.get(r["name"].strip().upper(), [])
        active = next((p for p in params if p.get("active")), None)
        x_name = (active or {}).get("x") or "Vehicle Speed"
        y_name = (active or {}).get("y") or "AccelerationChassis"
        pts = []
        for ev in events:
            x = get_channel(ev["channels"], x_name)
            y = get_channel(ev["channels"], y_name)
            try:
                pts.append({"x": float(x), "y": float(y),
                            "color": (ev.get("driv") or {}).get("color")})
            except (TypeError, ValueError):
                continue
        png_buf = report_builder.scatter_png(pts, x_name, y_name, r["name"])
        charts[r["name"]] = {"png": png_buf.getvalue() if png_buf else None}
    data = {"project": project, "global": glob, "sdv_results": rows,
            "charts": charts,
            "doc_versions": payload.get("doc_versions") or []}
    out = os.path.join(tempfile.gettempdir(), f"ODRIV_report.{fmt}")
    if fmt == "pptx":
        report_builder.build_pptx(data, out)
        media = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
    else:
        report_builder.build_pdf(data, out)
        media = "application/pdf"
    await moniteur(f"Report generated ({fmt.upper()})")
    fname = f"ODRIV_{(project.get('name_code') or 'report').replace(' ', '_')}.{fmt}"
    return FileResponse(out, filename=fname, media_type=media)


app.include_router(api)
app.add_middleware(
    CORSMiddleware,
    allow_credentials=True,
    allow_origins=os.environ.get('CORS_ORIGINS', '*').split(','),
    allow_methods=["*"],
    allow_headers=["*"],
)


@app.on_event("shutdown")
async def shutdown_db_client():
    client.close()
