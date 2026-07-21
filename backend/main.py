import json
import os
import threading
from contextlib import asynccontextmanager

from fastapi import FastAPI, Header, HTTPException, Query
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel


STATE_FILE = os.environ.get("STATE_FILE", "/data/state.json")
API_KEY = os.environ.get("API_KEY", "GT2026")
_lock = threading.Lock()


def _ensure_dir(path: str) -> None:
    directory = os.path.dirname(path)
    if directory:
        os.makedirs(directory, exist_ok=True)


def _load_state() -> dict:
    if not os.path.exists(STATE_FILE):
        return {}
    try:
        with open(STATE_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except (json.JSONDecodeError, OSError):
        return {}


def _save_state(data: dict) -> None:
    _ensure_dir(STATE_FILE)
    temp_path = STATE_FILE + ".tmp"
    with _lock:
        with open(temp_path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        os.replace(temp_path, STATE_FILE)


@asynccontextmanager
async def lifespan(app: FastAPI):
    _ensure_dir(STATE_FILE)
    _load_state()
    yield


app = FastAPI(
    title="GT Player Dashboard State",
    description="Shared checkbox state for promote/demote actions",
    lifespan=lifespan,
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)


class StateUpdate(BaseModel):
    week: str
    player: str
    status: str


@app.get("/health")
async def health():
    return {"ok": True}


@app.get("/state")
async def get_state(week: str = Query(..., description="Week label")):
    data = _load_state()
    week_data = data.get(week, {})
    return {"week": week, "states": week_data}


@app.post("/state")
async def post_state(payload: StateUpdate, x_api_key: str = Header(default="")):
    if x_api_key != API_KEY:
        raise HTTPException(status_code=403, detail="Invalid API key")
    data = _load_state()
    if payload.week not in data:
        data[payload.week] = {}
    data[payload.week][payload.player] = payload.status
    _save_state(data)
    return {"ok": True}


@app.get("/all")
async def get_all_states():
    return _load_state()
