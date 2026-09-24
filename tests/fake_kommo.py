"""API de Kommo simulada (sin red) y datos de prueba ficticios."""
from datetime import datetime, timezone, timedelta

TZ = timezone(timedelta(hours=-6))
CIERRES, VENTAS, OTRO = 1001, 2002, 3003


def ts(dia, hora=10, minuto=0):
    return int(datetime(2026, 9, dia, hora, minuto, tzinfo=TZ).timestamp())


def pipelines(nombres_142_143):
    """nombres_142_143: dict pipeline_id -> (nombre142, nombre143)."""
    def cerr(pid):
        a, b = nombres_142_143[pid]
        return [{"id": 142, "name": a, "sort": 10000}, {"id": 143, "name": b, "sort": 11000}]
    return {"_embedded": {"pipelines": [
        {"id": CIERRES, "name": "CIERRES", "_embedded": {"statuses": [
            {"id": 11, "name": "Contactado", "sort": 10},
            {"id": 12, "name": "Oferta", "sort": 20},
            {"id": 13, "name": "Oferta en espera", "sort": 30},
            {"id": 14, "name": "Documentación", "sort": 40},
            {"id": 15, "name": "Capturado", "sort": 50},
        ] + cerr(CIERRES)}},
        {"id": VENTAS, "name": "VENTAS", "_embedded": {"statuses": [
            {"id": 21, "name": "Nuevo", "sort": 10},
            {"id": 22, "name": "Contactado", "sort": 20},
            {"id": 23, "name": "SIN CAPACIDAD", "sort": 30},
            {"id": 24, "name": "BIMESTRE MAR-ABR 2026", "sort": 40},
            {"id": 25, "name": "Seguimiento 20 días", "sort": 50},
        ] + cerr(VENTAS)}},
        {"id": OTRO, "name": "OTRO", "_embedded": {"statuses": [
            {"id": 31, "name": "Entrada", "sort": 10},
        ] + cerr(OTRO)}},
    ]}}


NOMBRES_IGUALES = {p: ("Logrado con éxito", "Venta perdida") for p in (CIERRES, VENTAS, OTRO)}
NOMBRES_DISTINTOS = {CIERRES: ("Logrado con éxito", "Venta perdida"),
                     VENTAS: ("Ganado ventas", "Ventas perdidas"),
                     OTRO: ("Cerrado OTRO", "Perdido OTRO")}

_eid = [0]


def ev(lead, tipo, cuando, antes=None, despues=None, user=500):
    _eid[0] += 1
    def blk(x):
        if x is None:
            return []
        pid, sid = x
        return [{"lead_status": {"id": sid, "pipeline_id": pid}}]
    return {"id": f"e{_eid[0]}", "type": tipo, "entity_id": lead, "entity_type": "lead",
            "created_at": cuando, "created_by": user,
            "value_before": blk(antes), "value_after": blk(despues)}


def add(lead, cuando, pid, sid, user=500):
    return ev(lead, "lead_added", cuando, None, (pid, sid), user)


def mv(lead, cuando, de, a, user=500):
    return ev(lead, "lead_status_changed", cuando, de, a, user)


C, V, O = CIERRES, VENTAS, OTRO
EVENTOS = [
    # L101: CIERRES completo hasta ganado
    add(101, ts(2, 9), C, 11), mv(101, ts(5), (C, 11), (C, 12)),
    mv(101, ts(10), (C, 12), (C, 14)), mv(101, ts(15, 11, 30), (C, 14), (C, 142), user=501),
    # L102: CIERRES perdido
    add(102, ts(3), C, 11), mv(102, ts(8), (C, 11), (C, 12)), mv(102, ts(12), (C, 12), (C, 143)),
    # L103: VENTAS -> SIN CAPACIDAD (hoy ahí)
    add(103, ts(1, 0, 30), V, 21), mv(103, ts(4), (V, 21), (V, 22)), mv(103, ts(9), (V, 22), (V, 23)),
    # L104: VENTAS -> BIMESTRE 2026 (hoy ahí)
    add(104, ts(2), V, 21), mv(104, ts(11), (V, 21), (V, 24)),
    # L105: VENTAS -> CIERRES
    add(105, ts(2, 12), V, 21), mv(105, ts(6), (V, 21), (V, 22)),
    mv(105, ts(7), (V, 22), (C, 11)), mv(105, ts(14), (C, 11), (C, 13)),
    # L106: CIERRES -> VENTAS Contactado (etapa no filtrada)
    add(106, ts(5), C, 11), mv(106, ts(9, 16), (C, 11), (C, 12)), mv(106, ts(16), (C, 12), (V, 22)),
    # L107: sin FECHA DICTAMEN -> descartado
    add(107, ts(6), C, 11), mv(107, ts(8), (C, 11), (C, 15)),
    # L108: VENTAS "Seguimiento 20 días" (hoy ahí)
    add(108, ts(7), V, 21), mv(108, ts(13), (V, 21), (V, 25)),
    # L109: embudo OTRO, se ignora
    add(109, ts(4), O, 31), mv(109, ts(6), (O, 31), (O, 142)),
    # L111: CIERRES capturado, movimiento el día 20 a las 23:59
    add(111, ts(10), C, 11), mv(111, ts(20, 23, 59), (C, 11), (C, 15)),
    # L112: sin etiquetas
    add(112, ts(11), C, 11), mv(112, ts(18), (C, 11), (C, 13)),
    # L113: VENTAS en SIN CAPACIDAD pero con cambio fuera de rango (agosto)
    mv(113, int(datetime(2026, 8, 20, 10, tzinfo=TZ).timestamp()), (V, 21), (V, 23)),
    add(113, ts(12), V, 21), mv(113, ts(19), (V, 21), (V, 23)),
]

ACTUAL = {101: (C, 142), 102: (C, 143), 103: (V, 23), 104: (V, 24), 105: (C, 13),
          106: (V, 22), 107: (C, 15), 108: (V, 25), 109: (O, 142), 111: (C, 15),
          112: (C, 13), 113: (V, 23)}


def cf(nombre, tipo, valores):
    return {"field_id": hash(nombre) % 1000, "field_name": nombre, "field_type": tipo,
            "values": [{"value": v} for v in valores]}


CAMPOS = {
    101: [cf("Estatus de Negocio", "select", ["DICTAMINADO"]), cf("FECHA DICTAMEN", "date", [ts(14, 0)])],
    102: [cf("ESTATUS DE NEGOCIO", "select", ["RECHAZADO"]), cf("Fecha Dictámen", "date", [ts(11, 0)])],
    103: [cf("ESTATUS DE NEGOCIO", "multiselect", ["A", "B"]), cf("FECHA DICTAMEN", "date", [ts(8, 0)])],
    104: [cf("ESTATUS DE NEGOCIO", "select", ["PENDIENTE"]), cf("FECHA DICTAMEN", "date_time", [ts(10, 15, 45)])],
    105: [cf("ESTATUS DE NEGOCIO", "text", ["OK"]), cf("FECHA DICTAMEN", "date", [ts(13, 0)])],
    106: [cf("ESTATUS DE NEGOCIO", "select", ["DICTAMINADO"]), cf("FECHA DICTAMEN", "date", [ts(9, 0)])],
    107: [cf("ESTATUS DE NEGOCIO", "select", ["DICTAMINADO"]), cf("FECHA DICTAMEN", "date", [])],
    108: [cf("ESTATUS DE NEGOCIO", "checkbox", [True]), cf("FECHA DICTAMEN", "date", [ts(12, 0)])],
    109: [cf("ESTATUS DE NEGOCIO", "select", ["X"]), cf("FECHA DICTAMEN", "date", [ts(5, 0)])],
    111: [cf("ESTATUS DE NEGOCIO", "numeric", [42]), cf("FECHA DICTAMEN", "date", [ts(19, 0)])],
    112: [cf("ESTATUS DE NEGOCIO", "select", ["DICTAMINADO"]), cf("FECHA DICTAMEN", "date", [ts(17, 0)])],
    113: [cf("ESTATUS DE NEGOCIO", "select", ["DICTAMINADO"]), cf("FECHA DICTAMEN", "date", [ts(18, 0)])],
}
TAGS = {101: ["ASESOR UNO CIERRES"], 102: ["ANA PRUEBA CIERRES"], 103: ["JULIANA PRUEBA CIERRES"],
        104: ["ASESOR UNO CIERRES", "VIP"], 105: ["ANA PRUEBA CIERRES"], 106: ["ASESOR UNO CIERRES"],
        107: ["ASESOR UNO CIERRES"], 108: ["JULIANA PRUEBA CIERRES"], 109: ["ANA PRUEBA CIERRES"],
        111: ["ASESOR UNO CIERRES"], 112: [], 113: ["Ana Prueba Cierres"]}


def lead_obj(lid):
    pid, sid = ACTUAL[lid]
    return {"id": lid, "pipeline_id": pid, "status_id": sid,
            "custom_fields_values": CAMPOS.get(lid) or None,
            "_embedded": {"tags": [{"id": i, "name": n} for i, n in enumerate(TAGS.get(lid, []))]}}


class FakeResponse:
    def __init__(self, status, data=None, headers=None):
        self.status_code = status
        self._data = data
        self.headers = headers or {}

    def json(self):
        return self._data

    def raise_for_status(self):
        if self.status_code >= 400:
            import requests
            raise requests.exceptions.HTTPError(f"HTTP {self.status_code}")


def _pagina(items, page, size, key):
    ini = (page - 1) * size
    trozo = items[ini:ini + size]
    if not trozo:
        return FakeResponse(204)
    data = {"_embedded": {key: trozo}, "_links": {}}
    if ini + size < len(items):
        data["_links"]["next"] = {"href": "x"}
    return FakeResponse(200, data)


class FakeKommo:
    def __init__(self, nombres=NOMBRES_IGUALES, fallos_429=1):
        self.nombres = nombres
        self.llamadas = []
        self.fallos_429 = fallos_429   # primeras N llamadas a /events devuelven 429

    def get(self, url, params=None, timeout=None):
        params = dict(params or {})
        self.llamadas.append((url, params))
        ruta = url.split("/api/v4", 1)[1]
        if ruta == "/leads/pipelines":
            return FakeResponse(200, pipelines(self.nombres))
        if ruta == "/users":
            return _pagina([{"id": 500, "name": "Usuario Admin"}, {"id": 501, "name": "Usuario Dos"}],
                           params.get("page", 1), 1, "users")
        if ruta == "/events":
            if self.fallos_429 > 0:
                self.fallos_429 -= 1
                return FakeResponse(429, headers={"Retry-After": "0"})
            tipos = {v for k, v in params.items() if k.startswith("filter[type]")}
            d = params.get("filter[created_at][from]")
            h = params.get("filter[created_at][to]")
            items = [e for e in EVENTOS if e["type"] in tipos
                     and (d is None or e["created_at"] >= d)
                     and (h is None or e["created_at"] <= h)]
            items.sort(key=lambda e: -e["created_at"])
            return _pagina(items, params["page"], 3, "events")
        if ruta == "/leads":
            if "filter[id][]" in params:
                ids = params["filter[id][]"]
                items = [lead_obj(i) for i in ids if i in ACTUAL]
                return _pagina(items, params.get("page", 1), 4, "leads")
            pares = set()
            i = 0
            while f"filter[statuses][{i}][pipeline_id]" in params:
                pares.add((params[f"filter[statuses][{i}][pipeline_id]"],
                           params[f"filter[statuses][{i}][status_id]"]))
                i += 1
            items = [lead_obj(l) for l, par in sorted(ACTUAL.items()) if par in pares]
            return _pagina(items, params.get("page", 1), 2, "leads")
        raise AssertionError(f"URL inesperada: {url}")
