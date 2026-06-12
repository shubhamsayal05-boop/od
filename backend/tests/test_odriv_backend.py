"""End-to-end backend tests for ODRIV FastAPI port."""
import io
import os
import pytest
import requests

BASE_URL = os.environ['REACT_APP_BACKEND_URL'].rstrip('/') if os.environ.get('REACT_APP_BACKEND_URL') \
    else open('/app/frontend/.env').read().split('REACT_APP_BACKEND_URL=')[1].split('\n')[0].strip()
API = f"{BASE_URL}/api"


@pytest.fixture(scope="module")
def session():
    s = requests.Session()
    s.headers.update({"Content-Type": "application/json"})
    return s


# ------------------------------------------------------------- state / config
class TestStateAndConfig:
    def test_state_ok(self, session):
        r = session.get(f"{API}/state", timeout=30)
        assert r.status_code == 200, r.text
        data = r.json()
        assert "version" in data and "Python" in data["version"]
        assert "sheets" in data and "total_events" in data

    def test_config_sections_present(self, session):
        r = session.get(f"{API}/config", timeout=30)
        assert r.status_code == 200
        cfg = r.json()
        for key in ["catalog", "definitions", "structure", "targets",
                    "settings_blocks", "target_vehicle", "chart_params",
                    "criticity", "thresholds"]:
            assert key in cfg, f"missing config section: {key}"

    def test_config_get_section(self, session):
        r = session.get(f"{API}/config/catalog", timeout=15)
        assert r.status_code == 200
        cat = r.json()
        assert isinstance(cat, list) and len(cat) >= 60  # 64 SDVs

    def test_unknown_config_section_404(self, session):
        r = session.get(f"{API}/config/__nope__", timeout=10)
        assert r.status_code == 404


# ------------------------------------------------------------- project lifecycle
class TestProjectPipeline:
    PROJECT = {
        "name_code": "TEST_PYTEST_PROJECT",
        "mode": "AUTO",
        "fuel": "Diesel",
        "gears": "8AT",
        "software_milestone": "MS3",
        "priority": "P1",
        "version": "4.6",
        "odriv_milestone": "MS3",
        "area": "EU",
        "target_vehicle": "REF",
        "number_of_gears": 8,
    }

    def test_01_new_project(self, session):
        r = session.post(f"{API}/project/new", json=self.PROJECT, timeout=15)
        assert r.status_code == 200, r.text
        p = r.json()
        assert p["name_code"] == self.PROJECT["name_code"]
        assert p["fuel"] == "Diesel"
        assert "id" in p

    def test_02_update_project(self, session):
        r = session.put(f"{API}/project", json={"area": "NA"}, timeout=15)
        assert r.status_code == 200
        p = r.json()
        assert p["area"] == "NA"
        # other fields preserved
        assert p["name_code"] == self.PROJECT["name_code"]

    def test_03_import_demo(self, session):
        r = session.post(f"{API}/import/demo", timeout=60)
        assert r.status_code == 200, r.text
        data = r.json()
        assert data["imported"] > 400, f"too few events: {data}"
        assert data["classified"] == data["imported"], "all demo events should classify"
        assert len(data["per_sdv"]) >= 50, f"only {len(data['per_sdv'])} SDVs populated"

    def test_04_state_reflects_events(self, session):
        r = session.get(f"{API}/state", timeout=15)
        data = r.json()
        assert data["total_events"] > 400
        assert data["project"]["name_code"] == self.PROJECT["name_code"]
        assert len(data["sheets"]) >= 40

    def test_05_calculate_rating(self, session):
        r = session.post(f"{API}/rating/calculate", timeout=120)
        assert r.status_code == 200, r.text
        data = r.json()
        glob = data["global"]
        assert "driv" in glob and "dyn" in glob
        for k in ["driv", "dyn"]:
            assert "index" in glob[k]
            assert "rate_low" in glob[k]
            assert glob[k]["verdict"] in {"Low Risk", "Medium Risk", "High Risk"}
        assert data["sdv_count"] >= 50

    def test_06_get_rating(self, session):
        r = session.get(f"{API}/rating", timeout=15)
        assert r.status_code == 200
        data = r.json()
        assert data["global"] is not None
        assert len(data["rows"]) >= 50
        # rows have expected structure
        row = data["rows"][0]
        assert "name" in row and "driv" in row
        if row.get("driv"):
            assert "index" in row["driv"]

    def test_07_sdv_detail(self, session):
        r = session.get(f"{API}/rating", timeout=15)
        rows = r.json()["rows"]
        first = rows[0]["name"]
        r2 = session.get(f"{API}/sdv/{first}", timeout=15)
        assert r2.status_code == 200, r2.text
        d = r2.json()
        assert d["name"]
        assert "events" in d and len(d["events"]) > 0
        assert "structure" in d and "charts" in d and "targets" in d

    def test_08_sdv_detail_unknown(self, session):
        r = session.get(f"{API}/sdv/__doesnotexist__", timeout=10)
        assert r.status_code == 404

    def test_09_events_listing(self, session):
        r = session.get(f"{API}/events?limit=10", timeout=15)
        assert r.status_code == 200
        data = r.json()
        assert data["total"] > 400
        assert len(data["events"]) == 10
        assert "id" in data["events"][0]

    def test_10_event_update_and_delete(self, session):
        r = session.get(f"{API}/events?limit=1", timeout=10)
        ev = r.json()["events"][0]
        eid = ev["id"]
        new_chans = dict(ev["channels"])
        # add a custom field
        new_chans["TEST_FIELD"] = {"min": 1, "max": 2, "mean": 1.5, "value": 1.5}
        r2 = session.put(f"{API}/events/{eid}", json={"channels": new_chans}, timeout=15)
        assert r2.status_code == 200
        updated = r2.json()
        assert "TEST_FIELD" in updated["channels"]
        # delete
        r3 = session.delete(f"{API}/events/{eid}", timeout=10)
        assert r3.status_code == 200
        # verify gone
        r4 = session.delete(f"{API}/events/{eid}", timeout=10)
        assert r4.status_code == 404

    def test_11_set_as_target(self, session):
        r = session.post(f"{API}/rating/set-as-target", timeout=20)
        assert r.status_code == 200
        data = r.json()
        assert data["ok"] is True
        assert data["rows"] >= 50
        # verify target_vehicle now has matching vehicle
        r2 = session.get(f"{API}/config/target_vehicle", timeout=10)
        tv = r2.json()
        assert any(row.get("vehicle") == "TEST_PYTEST_PROJECT" for row in tv["rows"])

    def test_12_download_sample_xlsx(self, session):
        r = session.get(f"{API}/import/sample", timeout=60)
        assert r.status_code == 200
        ctype = r.headers.get("content-type", "")
        assert "spreadsheetml" in ctype or "officedocument" in ctype
        assert len(r.content) > 5000
        assert r.content[:2] == b"PK"  # xlsx is a zip

    def test_13_reimport_via_file_endpoint(self, session):
        # download sample, then re-upload
        s2 = requests.Session()  # no Content-Type
        d = s2.get(f"{API}/import/sample", timeout=60)
        # need to wipe events first, but new project deletes all - skip wipe
        # Just verify upload works
        files = {"file": ("sample.xlsx", io.BytesIO(d.content),
                          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")}
        r = requests.post(f"{API}/import/file", files=files, timeout=60)
        assert r.status_code == 200, r.text
        data = r.json()
        assert data["imported"] > 400
        assert data["classified"] == data["imported"]

    def test_14_report_pptx(self, session):
        # Need fresh rating since reimport added more events
        rc = session.post(f"{API}/rating/calculate", timeout=120)
        assert rc.status_code == 200
        r = session.post(f"{API}/report/pptx",
                         json={"doc_versions": [{"name": "Test", "version": "1.0"}]},
                         timeout=180)
        assert r.status_code == 200, r.text
        assert len(r.content) > 5000
        assert r.content[:2] == b"PK"  # pptx is a zip

    def test_15_report_pdf(self, session):
        r = session.post(f"{API}/report/pdf", json={}, timeout=180)
        assert r.status_code == 200, r.text
        assert len(r.content) > 1000
        assert r.content[:4] == b"%PDF"

    def test_16_report_bad_fmt(self, session):
        r = session.post(f"{API}/report/docx", json={}, timeout=10)
        assert r.status_code == 400

    def test_17_logs(self, session):
        r = session.get(f"{API}/logs?limit=20", timeout=10)
        assert r.status_code == 200
        logs = r.json()
        assert isinstance(logs, list) and len(logs) > 0
        assert all("ts" in l and "message" in l for l in logs)


# ------------------------------------------------------------- config edit roundtrip
class TestConfigEditRoundtrip:
    def test_targets_put_and_persist(self, session):
        r = session.get(f"{API}/config/targets", timeout=10)
        assert r.status_code == 200
        data = r.json()
        # PUT it back as-is
        r2 = session.put(f"{API}/config/targets", json={"data": data}, timeout=15)
        assert r2.status_code == 200
        assert r2.json()["ok"] is True
        # GET back
        r3 = session.get(f"{API}/config/targets", timeout=10)
        assert r3.status_code == 200
        # structural equality (key set)
        assert set(r3.json().keys()) == set(data.keys())

    def test_put_unknown_section(self, session):
        r = session.put(f"{API}/config/__nope__", json={"data": {}}, timeout=10)
        assert r.status_code == 404


# ------------------------------------------------------------- cleanup
class TestCleanup:
    def test_zz_erase_all(self, session):
        r = session.delete(f"{API}/project", timeout=15)
        assert r.status_code == 200
        st = session.get(f"{API}/state", timeout=10).json()
        assert st["project"] is None
        assert st["total_events"] == 0
        assert st["has_rating"] is False
