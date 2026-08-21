#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
成都看雪山 · Flask 单文件版 v1.8（观山季点位版）

安装：
    python -m pip install flask requests
运行：
    python chengdu_snow_mountain.py
浏览器：
    http://127.0.0.1:5000

气溶胶来自 Open-Meteo Air Quality API 的 aerosol_optical_depth（550 nm）。
无需 API Key、NetCDF、xarray、cdsapi 或 ecCodes。
"""

from __future__ import annotations

import math
import os
import threading
import time
from datetime import datetime, timedelta, timezone

import requests
from flask import Flask, jsonify, request, render_template_string

APP = Flask(__name__)
TZ = timezone(timedelta(hours=8))
DEFAULT_OBSERVER = {"name": "成都", "lat": 30.65957, "lon": 104.05511, "elev": 500}
MOUNTAINS = [
    {"id": "xiling", "name": "西岭雪山·大雪塘", "lat": 30.6167, "lon": 103.1700, "elev": 5364},
    {"id": "siguniang", "name": "四姑娘山·幺妹峰", "lat": 31.109766, "lon": 102.905291, "elev": 6250},
    {"id": "gongga", "name": "贡嘎山主峰", "lat": 29.595833, "lon": 101.879167, "elev": 7508},
    {"id": "jiuding", "name": "九顶山·狮子王峰", "lat": 31.6470, "lon": 103.8270, "elev": 4984},
]
VIEWPOINTS = [
    {"name":"凤凰山公园至真观","district":"金牛区","lat":30.737724,"lon":104.081896,"precision":"园区坐标","target":"西北—西侧群峰"},
    {"name":"成都大熊猫繁育研究基地","district":"成华区","lat":30.742306,"lon":104.135901,"precision":"园区坐标","target":"西北—西侧群峰"},
    {"name":"锦城湖3号湖区广场","district":"高新区","lat":30.573865,"lon":104.039910,"precision":"湖区近似位置","target":"西岭、幺妹峰"},
    {"name":"钟家山美满村三号观景平台","district":"龙泉驿区","lat":30.5620,"lon":104.3300,"precision":"景区近似位置","target":"成都平原与西部群峰"},
    {"name":"丹景山狮子堡观景台","district":"四川天府新区","lat":30.3900,"lon":104.2100,"precision":"景区近似位置","target":"西岭雪山、城市天际线"},
    {"name":"鲁家滩湿地公园","district":"温江区","lat":30.689937,"lon":103.774798,"precision":"园区坐标","target":"西岭与龙门山群峰"},
    {"name":"世界科幻公园·菁蓉湖","district":"郫都区","lat":30.8060,"lon":103.8780,"precision":"园区近似位置","target":"龙门山群峰"},
    {"name":"天府农博岛","district":"新津区","lat":30.504319,"lon":103.801021,"precision":"园区坐标","target":"西岭、贡嘎方向"},
    {"name":"都江堰南桥","district":"都江堰市","lat":30.997597,"lon":103.616407,"precision":"街区坐标","target":"龙门山近景"},
    {"name":"长秋月观景台","district":"蒲江县","lat":30.2500,"lon":103.6000,"precision":"长秋山近似位置","target":"贡嘎、幺妹峰及川西群峰"},
]
WEATHER_MODELS = {
    "best_match": {"name": "智能最佳匹配", "detail": "Open-Meteo 自动选择当地最佳可用模型"},
    "ecmwf_ifs025": {"name": "ECMWF IFS 0.25°", "detail": "欧洲中期天气预报中心全球模式"},
    "gfs_seamless": {"name": "NOAA GFS", "detail": "美国 NOAA 全球预报系统"},
    "icon_seamless": {"name": "DWD ICON", "detail": "德国气象局全球模式"},
    "cma_grapes_global": {"name": "CMA GRAPES", "detail": "中国气象局约15 km全球模式"},
    "jma_seamless": {"name": "JMA", "detail": "日本气象厅全球模式"},
}

SESSION = requests.Session()
SESSION.headers.update({"User-Agent": "ChengduSnowMountain/1.0"})
_cache = {}
_cache_lock = threading.Lock()


def clamp(x, lo=0.0, hi=1.0):
    return max(lo, min(hi, x))


def haversine_bearing(a_lat, a_lon, b_lat, b_lon):
    p1, p2 = map(math.radians, (a_lat, b_lat))
    dl = math.radians(b_lon - a_lon)
    y = math.sin(dl) * math.cos(p2)
    x = math.cos(p1) * math.sin(p2) - math.sin(p1) * math.cos(p2) * math.cos(dl)
    bearing = (math.degrees(math.atan2(y, x)) + 360) % 360
    h = math.sin((p2-p1)/2)**2 + math.cos(p1)*math.cos(p2)*math.sin(dl/2)**2
    return 6371.0088 * 2 * math.asin(math.sqrt(h)), bearing


def interpolate_great_circle(lat1, lon1, lat2, lon2, n=10):
    # 距离小于 300 km 时线性插值误差对 0.4°/0.25°气象格点可忽略。
    return [(lat1+(lat2-lat1)*i/(n-1), lon1+(lon2-lon1)*i/(n-1)) for i in range(n)]


def solar_position(dt_local, lat, lon):
    """NOAA 近似太阳高度角/方位角；dt_local 必须带时区。"""
    u = dt_local.astimezone(timezone.utc)
    jd = u.timestamp()/86400.0 + 2440587.5
    t = (jd-2451545.0)/36525.0
    L0 = (280.46646 + t*(36000.76983 + 0.0003032*t)) % 360
    M = math.radians((357.52911 + t*(35999.05029-0.0001537*t)) % 360)
    C = (1.914602-t*(0.004817+0.000014*t))*math.sin(M) + (0.019993-0.000101*t)*math.sin(2*M) + 0.000289*math.sin(3*M)
    lam = math.radians(L0+C-0.00569-0.00478*math.sin(math.radians(125.04-1934.136*t)))
    eps = math.radians(23.439291-0.0130042*t)
    dec = math.asin(math.sin(eps)*math.sin(lam))
    y = math.tan(eps/2)**2
    L = math.radians(L0)
    eq = 4*math.degrees(y*math.sin(2*L)-2*0.016708634*math.sin(M)+4*0.016708634*y*math.sin(M)*math.cos(2*L)-0.5*y*y*math.sin(4*L)-1.25*0.016708634**2*math.sin(2*M))
    minutes = u.hour*60+u.minute+u.second/60
    tst = (minutes + eq + 4*lon) % 1440
    ha = math.radians(tst/4-180)
    phi = math.radians(lat)
    zen = math.acos(clamp(math.sin(phi)*math.sin(dec)+math.cos(phi)*math.cos(dec)*math.cos(ha), -1, 1))
    elev = 90-math.degrees(zen)
    az = (math.degrees(math.atan2(math.sin(ha), math.cos(ha)*math.sin(phi)-math.tan(dec)*math.cos(phi)))+180) % 360
    return elev, az


def apparent_peak_angle(distance_km, peak_elev, observer_elev):
    # 标准折射取地球等效半径 7/6 R，得到峰顶几何仰角。
    re = 6371008.8 * 7/6
    d = distance_km*1000
    drop = d*d/(2*re)
    return math.degrees(math.atan2(peak_elev-observer_elev-drop, d))


def cache_get(key, max_age):
    with _cache_lock:
        item = _cache.get(key)
        return item[1] if item and time.time()-item[0] < max_age else None


def cache_put(key, value):
    with _cache_lock:
        _cache[key] = (time.time(), value)


def open_meteo_elevations(points):
    """Open-Meteo GLO-90 海拔；官方单次最多 100 个坐标。"""
    values = []
    for start in range(0, len(points), 100):
        batch = points[start:start+100]
        params = {
            "latitude": ",".join(f"{p[0]:.6f}" for p in batch),
            "longitude": ",".join(f"{p[1]:.6f}" for p in batch),
        }
        r = SESSION.get("https://api.open-meteo.com/v1/elevation", params=params, timeout=30)
        r.raise_for_status()
        part = r.json().get("elevation", [])
        if len(part) != len(batch):
            raise RuntimeError(f"海拔 API 返回 {len(part)} 个值，预期 {len(batch)} 个")
        values.extend(None if x is None else float(x) for x in part)
    return values


def observer_elevation(lat, lon):
    key = f"observer-elev:{lat:.5f},{lon:.5f}"
    cached = cache_get(key, 86400 * 30)
    if cached is not None: return cached
    value = open_meteo_elevations([(lat, lon)])[0]
    if value is None: raise RuntimeError("Open-Meteo 未返回观测点海拔")
    cache_put(key, value)
    return value


def path_terrain(observer, mountain, n=10):
    points=interpolate_great_circle(observer["lat"],observer["lon"],mountain["lat"],mountain["lon"],n)
    key="path-terrain:"+":".join(f"{a:.4f},{b:.4f}" for a,b in points)
    cached=cache_get(key,86400*30)
    if cached is not None: return cached
    values=open_meteo_elevations(points)
    # 端点使用已知观测点/峰顶高程，避免90m DEM格点把尖锐峰顶平滑掉。
    values[0]=float(observer["elev"]); values[-1]=float(mountain["elev"])
    cache_put(key,values)
    return values


def open_meteo_corridor(observer, mountains, days=5, model="best_match"):
    points, owners = [], []
    for m in mountains:
        for lat, lon in interpolate_great_circle(observer["lat"], observer["lon"], m["lat"], m["lon"], 10):
            points.append((lat, lon)); owners.append(m["id"])
    key = "om:" + model + ":" + ":".join(f"{x[0]:.3f},{x[1]:.3f}" for x in points)
    cached = cache_get(key, 1800)
    if cached: return cached
    params = {
        "latitude": ",".join(f"{p[0]:.4f}" for p in points),
        "longitude": ",".join(f"{p[1]:.4f}" for p in points),
        "hourly": "cloud_cover,cloud_cover_low,cloud_cover_mid,cloud_cover_high,visibility,relative_humidity_2m,precipitation_probability,precipitation,wind_speed_10m,temperature_2m",
        "forecast_days": days, "timezone": "Asia/Shanghai", "models": model
    }
    r = SESSION.get("https://api.open-meteo.com/v1/forecast", params=params, timeout=30)
    r.raise_for_status()
    raw = r.json()
    rows = raw if isinstance(raw, list) else [raw]
    result = {m["id"]: [] for m in mountains}
    for i, row in enumerate(rows): result[owners[i]].append(row["hourly"])
    cache_put(key, result)
    return result


def open_meteo_aerosol(observer, mountains, days=5):
    """获取每座山视线上的 10 个 AOD550/PM2.5/沙尘采样点。"""
    points, owners = [], []
    for m in mountains:
        for lat, lon in interpolate_great_circle(observer["lat"], observer["lon"], m["lat"], m["lon"], 10):
            points.append((lat, lon)); owners.append(m["id"])
    key = "air:" + ":".join(f"{x[0]:.3f},{x[1]:.3f}" for x in points)
    cached = cache_get(key, 3600)
    if cached: return cached
    params = {
        "latitude": ",".join(f"{p[0]:.4f}" for p in points),
        "longitude": ",".join(f"{p[1]:.4f}" for p in points),
        "hourly": "aerosol_optical_depth,pm2_5,dust",
        "forecast_days": days,
        "timezone": "Asia/Shanghai",
        "domains": "cams_global",
        "cell_selection": "nearest",
    }
    r = SESSION.get("https://air-quality-api.open-meteo.com/v1/air-quality", params=params, timeout=35)
    r.raise_for_status()
    raw = r.json()
    rows = raw if isinstance(raw, list) else [raw]
    if len(rows) != len(points):
        raise RuntimeError(f"Open-Meteo 气溶胶返回 {len(rows)} 个点，预期 {len(points)} 个点")
    result = {m["id"]: [] for m in mountains}
    for i, row in enumerate(rows):
        result[owners[i]].append(row["hourly"])
    cache_put(key, result)
    return result


def air_at_time(corridor, local_dt):
    """返回视线最不利 AOD，以及用于解释的 PM2.5/沙尘峰值。"""
    if not corridor: return {"aod": None, "pm2_5": None, "dust": None, "points": []}
    times = corridor[0].get("time", [])
    if not times: return {"aod": None, "pm2_5": None, "dust": None, "points": []}
    idx = min(range(len(times)), key=lambda i: abs(datetime.fromisoformat(times[i]).replace(tzinfo=TZ).timestamp()-local_dt.timestamp()))
    def maximum(field):
        values=[]
        for point in corridor:
            arr=point.get(field, [])
            if idx < len(arr) and arr[idx] is not None:
                try: values.append(float(arr[idx]))
                except (TypeError, ValueError): pass
        return max(values) if values else None
    points=[]
    for point in corridor:
        item={}
        for field in ("aerosol_optical_depth","pm2_5","dust"):
            arr=point.get(field,[]); value=arr[idx] if idx<len(arr) else None
            try: item[field]=None if value is None else float(value)
            except (TypeError,ValueError): item[field]=None
        points.append(item)
    return {"aod": maximum("aerosol_optical_depth"), "pm2_5": maximum("pm2_5"), "dust": maximum("dust"), "points": points}


def score_hour(dt, mountain, corridor, observer, air, terrain):
    aod = air.get("aod")
    idx = min(range(len(corridor[0]["time"])), key=lambda i: abs(datetime.fromisoformat(corridor[0]["time"][i]).replace(tzinfo=TZ).timestamp()-dt.timestamp()))
    def vals(field, default=0):
        out=[]
        for p in corridor:
            arr=p.get(field) or []
            value=arr[idx] if idx < len(arr) and arr[idx] is not None else default
            try: out.append(float(value))
            except (TypeError,ValueError): out.append(float(default))
        return out
    low, mid, high = vals("cloud_cover_low"), vals("cloud_cover_mid"), vals("cloud_cover_high")
    vis = vals("visibility", 10000); rh = vals("relative_humidity_2m", 70); pop = vals("precipitation_probability")
    wind = vals("wind_speed_10m")
    prior_rain=[]
    for p in corridor:
        arr=p.get("precipitation") or []
        prior_rain.append(sum(float(v or 0) for v in arr[max(0,idx-12):idx]))
    # 近观测端低云/雾权重最大；山体附近中低云决定峰是否被包裹。
    cloud_pen = 0.34*max(low[:3])/100 + 0.28*max(low[3:])/100 + 0.17*max(mid[3:])/100 + 0.06*sum(high)/len(high)/100
    vis_factor = clamp((min(vis[:3])-5000)/45000)
    humidity_pen = clamp((max(rh[:3])-65)/35)*0.14
    rain_pen = max(pop)/100*0.12
    aerosol_factor = None if aod is None else math.exp(-2.15*max(0,aod))
    aerosol_pen = 0.18 if aerosol_factor is None else (1-aerosol_factor)*0.36
    base = clamp(1-cloud_pen-humidity_pen-rain_pen-aerosol_pen)
    # 超远山对能见度和气溶胶更敏感。
    dist, bearing = haversine_bearing(observer["lat"],observer["lon"],mountain["lat"],mountain["lon"])
    range_factor = clamp(vis_factor + 0.35, 0.18, 1) ** (dist/180)
    # 观山经验只作小幅修正，实时云量、能见度和AOD仍是主约束。
    washout_bonus = min(6.0, 2.2*max(prior_rain[:3])) if prior_rain else 0
    wind_bonus = 3.0 if 8 <= sum(wind[:3])/max(1,len(wind[:3])) <= 28 else 0
    season_bonus = 2.0 if 3 <= dt.month <= 9 else 0
    time_bonus = 2.0 if (5 <= dt.hour <= 8 or 18 <= dt.hour <= 19) else 0
    empirical_bonus = washout_bonus+wind_bonus+season_bonus+time_bonus
    score = round(clamp(100*base*range_factor+empirical_bonus,0,100))
    elev, sunaz = solar_position(dt, observer["lat"], observer["lon"])
    gold = 0
    if -1.5 <= elev <= 6:
        # 太阳位于山峰相反方向附近时，山体正面受光（容差随散射放宽）。
        sep = abs((sunaz-bearing+180)%360-180)
        front = clamp((sep-70)/70)
        gold = round(score*front*clamp((elev+1.5)/3.5 if elev<2 else (6-elev)/4+0.35))
    reasons=[]
    if max(low)>65: reasons.append("通道低云偏多")
    if max(mid[3:])>65: reasons.append("峰区中云遮挡")
    if min(vis[:3])<15000: reasons.append("盆地能见度不足")
    if aod is not None and aod>0.4: reasons.append("AOD550 很高")
    elif aod is not None and aod>0.2: reasons.append("气溶胶偏多")
    if max(rh[:3])>88: reasons.append("近地层湿度高")
    if washout_bonus>=2: reasons.append("前期降水有利于洗除气溶胶")
    if wind_bonus: reasons.append("近地风有利于盆地扩散")
    profile=[]
    air_points=air.get("points",[])
    for i in range(len(corridor)):
        ap=air_points[i] if i<len(air_points) else {}
        profile.append({"distance":round(dist*i/max(1,len(corridor)-1),1),"terrain":round(float(terrain[i]),1),"low":round(low[i],1),"mid":round(mid[i],1),"high":round(high[i],1),"rh":round(rh[i],1),"visibility":round(vis[i]/1000,1),"aod":ap.get("aerosol_optical_depth")})
    return {"time":dt.isoformat(),"score":score,"gold":gold,"empirical_bonus":round(empirical_bonus,1),"aod":None if aod is None else round(aod,3),"pm2_5":None if air.get("pm2_5") is None else round(air["pm2_5"],1),"dust":None if air.get("dust") is None else round(air["dust"],1),"low":round(max(low),1),"mid":round(max(mid[3:]),1),"high":round(sum(high)/len(high),1),"visibility":round(min(vis[:3])/1000,1),"rh":round(max(rh[:3]),1),"sun_elev":round(elev,1),"profile":profile,"reasons":reasons or ["云量与通透度较好"]}


def build_forecast(observer, model="best_match"):
    if model not in WEATHER_MODELS: model="best_match"
    meteo = open_meteo_corridor(observer, MOUNTAINS, model=model)
    aerosol = open_meteo_aerosol(observer, MOUNTAINS)
    result=[]
    for m in MOUNTAINS:
        dist,bearing=haversine_bearing(observer["lat"],observer["lon"],m["lat"],m["lon"])
        terrain=path_terrain(observer,m,len(meteo[m["id"]]))
        hours=[]
        # 每小时评价，页面展示每天清晨/傍晚最佳值。
        times=[datetime.fromisoformat(t).replace(tzinfo=TZ) for t in meteo[m["id"]][0]["time"]]
        for dt in times:
            air=air_at_time(aerosol[m["id"]],dt)
            hours.append(score_hour(dt,m,meteo[m["id"]],observer,air,terrain))
        daily=[]
        dates=sorted(set(x["time"][:10] for x in hours))
        for d in dates:
            day=[x for x in hours if x["time"].startswith(d)]
            morning=[x for x in day if 5<=datetime.fromisoformat(x["time"]).hour<=10]
            evening=[x for x in day if 16<=datetime.fromisoformat(x["time"]).hour<=20]
            best=max(day,key=lambda x:x["score"])
            daily.append({"date":d,"morning":max(morning,key=lambda x:x["score"]) if morning else best,"evening":max(evening,key=lambda x:x["score"]) if evening else best,"gold":max(day,key=lambda x:x["gold"])})
        # 前端只需每天晨/晚最佳时刻；不回传全部逐小时剖面，显著减少手机流量。
        result.append({**m,"distance":round(dist,1),"bearing":round(bearing,1),"peak_angle":round(apparent_peak_angle(dist,m["elev"],observer["elev"]),2),"daily":daily})
    return {"observer":observer,"mountains":result,"weather_model":{"id":model,**WEATHER_MODELS[model]},"aerosol":{"status":"ready","message":"Open-Meteo CAMS Global · AOD550/PM2.5/沙尘 · 自动缓存1小时"},"generated":datetime.now(TZ).isoformat(),"method":"所选天气模型 + Open-Meteo AOD550 + 通道云量/湿度/能见度 + 前期降水洗除 + 风场扩散 + 太阳几何"}


@APP.get("/")
def index(): return render_template_string(HTML, observer=DEFAULT_OBSERVER, viewpoints=VIEWPOINTS)


@APP.get("/api/forecast")
def api_forecast():
    try:
        lat=float(request.args.get("lat",DEFAULT_OBSERVER["lat"])); lon=float(request.args.get("lon",DEFAULT_OBSERVER["lon"]))
        model=request.args.get("model","best_match")
        if model not in WEATHER_MODELS: raise ValueError("不支持的天气数据源")
        observer={"name":request.args.get("name","成都"),"lat":lat,"lon":lon,"elev":round(observer_elevation(lat,lon),1)}
        return jsonify({"ok":True,"data":build_forecast(observer,model)})
    except Exception as e:
        return jsonify({"ok":False,"error":f"{type(e).__name__}: {e}"}),500


@APP.get("/api/elevation-grid")
def api_elevation_grid():
    try:
        south=float(request.args["south"]); west=float(request.args["west"])
        north=float(request.args["north"]); east=float(request.args["east"])
        n=max(6,min(14,int(request.args.get("n",12))))
        if not (-90 <= south < north <= 90 and -180 <= west < east <= 180):
            raise ValueError("地图范围无效")
        # 采用格心采样；每格返回边界，前端直接绘制半透明分区。
        dy=(north-south)/n; dx=(east-west)/n
        points=[(south+(iy+.5)*dy,west+(ix+.5)*dx) for iy in range(n) for ix in range(n)]
        key=f"elev-grid:{south:.3f},{west:.3f},{north:.3f},{east:.3f},{n}"
        elevs=cache_get(key,86400*30)
        if elevs is None:
            elevs=open_meteo_elevations(points); cache_put(key,elevs)
        cells=[]
        for k,((lat,lon),z) in enumerate(zip(points,elevs)):
            iy=k//n; ix=k%n
            band="unknown" if z is None else ("low" if z<1500 else "mid" if z<3500 else "high")
            cells.append({"south":south+iy*dy,"north":south+(iy+1)*dy,"west":west+ix*dx,"east":west+(ix+1)*dx,"elevation":z,"band":band})
        return jsonify({"ok":True,"cells":cells,"source":"Open-Meteo GLO-90","thresholds":{"low":"<1500 m","mid":"1500–3500 m","high":"≥3500 m"}})
    except Exception as e:
        return jsonify({"ok":False,"error":f"{type(e).__name__}: {e}"}),400


HTML = r'''<!doctype html><html lang="zh-CN"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover">
<title>成都看雪山 · Open-Meteo气溶胶版</title><link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css">
<style>
:root{--bg:#071019;--panel:#0d1b29;--line:#213a50;--text:#eaf6ff;--muted:#8faec2;--cyan:#43d9ff;--gold:#ffcc66;--red:#ff6b75;--green:#45e0a8}*{box-sizing:border-box}body{margin:0;background:radial-gradient(circle at 80% 0,#153957 0,#071019 42%);color:var(--text);font-family:system-ui,"Noto Sans SC",sans-serif}.wrap{max-width:1180px;margin:auto;padding:16px}.hero{display:flex;justify-content:space-between;gap:14px;align-items:end;margin:12px 0 18px}.hero h1{margin:0;font-size:clamp(26px,5vw,44px);letter-spacing:2px}.hero p{margin:6px 0;color:var(--muted)}button{border:1px solid #2a6d8c;background:#12354b;color:white;border-radius:10px;padding:11px 14px;font-weight:700}.primary{background:linear-gradient(135deg,#1a9ec8,#205cc2)}.grid{display:grid;grid-template-columns:340px 1fr;gap:14px}.panel{background:#0d1b29e8;border:1px solid var(--line);border-radius:16px;padding:14px;box-shadow:0 12px 30px #0005}.inputs{display:grid;grid-template-columns:1fr 1fr;gap:9px}.inputs label{font-size:12px;color:var(--muted)}.inputs .wide{grid-column:1/-1}input,select{width:100%;margin-top:4px;background:#07131e;border:1px solid #29465c;color:white;border-radius:9px;padding:10px}.actions{display:flex;gap:8px;margin-top:12px}.status{margin-top:12px;padding:10px;border-left:3px solid var(--cyan);background:#091622;color:var(--muted);font-size:13px;word-break:break-word}#map{height:340px;border-radius:12px}.legend{display:flex;gap:12px;flex-wrap:wrap;font-size:12px;color:var(--muted);margin-top:8px}.cards{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:12px;margin-top:14px}.mountain h3{margin:0 0 4px}.meta{font-size:12px;color:var(--muted)}.days{display:grid;grid-template-columns:repeat(5,minmax(116px,1fr));gap:7px;overflow:auto;margin-top:12px}.day{background:#081522;border:1px solid #19364b;border-radius:10px;padding:8px;min-width:116px}.day strong{font-size:12px}.score{font-size:26px;font-weight:900;margin:7px 0}.bar{height:5px;background:#1d3040;border-radius:5px;overflow:hidden}.bar i{display:block;height:100%;background:linear-gradient(90deg,var(--red),var(--gold),var(--green))}.small{font-size:11px;color:var(--muted);line-height:1.65}.tabs{display:flex;gap:4px;margin:6px 0}.tabs span{font-size:11px;padding:3px 6px;background:#132638;border-radius:6px}.sim-actions{display:flex;gap:4px;margin-top:7px}.sim-actions button{padding:6px 7px;font-size:10px;flex:1}.loading{opacity:.6;pointer-events:none}.foot{color:var(--muted);font-size:12px;line-height:1.7;margin:14px 0}.modal{position:fixed;inset:0;background:#000b;z-index:9999;display:none;align-items:center;justify-content:center;padding:12px}.modal.open{display:flex}.sim-box{width:min(900px,100%);max-height:96vh;overflow:auto;background:#091725;border:1px solid #31536d;border-radius:17px;padding:13px;box-shadow:0 25px 80px #000}.sim-head{display:flex;justify-content:space-between;align-items:start;gap:10px}.sim-head h3{margin:0}.sim-close{padding:7px 11px}.sim-canvas{width:100%;aspect-ratio:16/9;display:block;background:#19324a;border-radius:12px;margin-top:10px}.sim-metrics{display:grid;grid-template-columns:repeat(6,1fr);gap:6px;margin-top:8px}.sim-metrics div{background:#0e2232;border-radius:8px;padding:7px;text-align:center;font-size:11px;color:var(--muted)}.sim-metrics b{display:block;color:white;font-size:14px}.sim-note{font-size:12px;color:var(--muted);line-height:1.65;margin-top:8px}.evidence{margin-top:14px}.evidence summary{cursor:pointer;font-weight:800}.sample-list{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:7px;margin-top:10px}.sample{background:#081522;border:1px solid #19364b;border-radius:9px;padding:9px;font-size:12px}.sample a{color:#56d8ff;text-decoration:none}.tag{display:inline-block;margin-top:5px;padding:2px 6px;border-radius:9px;background:#16344a;color:#a9c9da;font-size:10px}@media(max-width:780px){.grid{grid-template-columns:1fr}.cards{grid-template-columns:1fr}.hero{align-items:start;flex-direction:column}#map{height:290px}.sim-metrics{grid-template-columns:repeat(3,1fr)}.sample-list{grid-template-columns:1fr}}
</style></head><body><div class="wrap"><div class="hero"><div><h1>成都 · 看雪山</h1><p>Open-Meteo AOD550 气溶胶 × 成都至峰区云廊 × 日照金山</p></div><button class="primary" onclick="loadAll()">刷新预报</button></div>
<div class="grid"><section class="panel"><div class="inputs"><label>纬度<input id="lat" value="{{observer.lat}}"></label><label>经度<input id="lon" value="{{observer.lon}}"></label><label>海拔 m（自动）<input id="elev" value="读取中" readonly></label><label>地点<input id="name" value="{{observer.name}}"></label><label class="wide">天气预报数据源<select id="model" onchange="loadAll()"><option value="best_match">智能最佳匹配（推荐）</option><option value="ecmwf_ifs025">ECMWF IFS 0.25°</option><option value="gfs_seamless">NOAA GFS</option><option value="icon_seamless">DWD ICON</option><option value="cma_grapes_global">中国气象局 CMA GRAPES</option><option value="jma_seamless">日本气象厅 JMA</option></select></label></div><div class="actions"><button onclick="locate()">手机定位</button><button onclick="loadAll()">更新全部数据</button></div><div id="modelStatus" class="status">天气模型：正在获取…</div><div id="aerosol" class="status">气溶胶：正在获取 Open-Meteo 数据…</div><div class="foot">天气模型可独立切换；AOD550 固定使用 Open-Meteo CAMS Global。观测点海拔和地图分区使用免费的 GLO-90 DEM。</div></section><section class="panel"><div id="map"></div><div class="legend"><span>● 观测点</span><span style="color:#ffcc66">● 雪山</span><span>线段 = 实际观测方向</span><span style="color:#54d68b">■ 低海拔 &lt;1500m</span><span style="color:#ffc857">■ 中海拔 1500–3500m</span><span style="color:#d88cff">■ 高海拔 ≥3500m</span></div></section></div>
<div id="cards" class="cards"></div><div class="foot">评分含义：80–100 极佳，65–79 较好，45–64 勉强，0–44 不推荐。AOD550 &lt; 0.10 很通透，0.10–0.20 尚好，0.20–0.40 偏灰，&gt; 0.40 对远山影响很大。评分已加入前12小时降水洗除、适度风速与观山季时段的小幅经验修正。</div></div>
<div id="simModal" class="modal" onclick="if(event.target===this)closeSimulation()"><div class="sim-box"><div class="sim-head"><div><h3 id="simTitle">天空形态模拟</h3><div id="simSub" class="meta"></div></div><button class="sim-close" onclick="closeSimulation()">关闭</button></div><canvas id="simCanvas" class="sim-canvas"></canvas><div id="simMetrics" class="sim-metrics"></div><div id="simNote" class="sim-note"></div></div></div>
<script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script><script>
const gaodeOpt={subdomains:'1234',maxZoom:18,attribution:'地图 © 高德'};
// 部分手机网络无法访问 webrd 域；标准图优先使用与卫星图相同、已验证可达的 webst 域。
const gaodeRoad=L.tileLayer('https://webst0{s}.is.autonavi.com/appmaptile?lang=zh_cn&size=1&scale=1&style=7&x={x}&y={y}&z={z}',gaodeOpt);
// 若当地节点不在 webst 提供 style=7，则逐瓦片回退到传统 webrd 标准图地址。
gaodeRoad.on('tileerror',e=>{const t=e.tile;if(t.dataset.gaodeRetry)return;t.dataset.gaodeRetry='1';t.src=t.src.replace(/webst0[1-4]/,'webrd02')});
const gaodeSat=L.tileLayer('https://webst0{s}.is.autonavi.com/appmaptile?style=6&x={x}&y={y}&z={z}',gaodeOpt);
const gaodeLabel=L.tileLayer('https://webst0{s}.is.autonavi.com/appmaptile?style=8&x={x}&y={y}&z={z}',gaodeOpt);
const gaodeSatellite=L.layerGroup([gaodeSat,gaodeLabel]);
const elevationLayer=L.layerGroup();
const viewpointLayer=L.layerGroup();
let map=L.map('map',{layers:[gaodeRoad,elevationLayer,viewpointLayer]}).setView([30.657,104.058],8),layers=[];
map.createPane('terrainPane');map.getPane('terrainPane').style.zIndex=350;
L.control.layers({'高德标准地图':gaodeRoad,'高德卫星影像':gaodeSatellite},{'成都观山季推荐点':viewpointLayer,'海拔分区':elevationLayer},{position:'topright',collapsed:true}).addTo(map);
const $=x=>document.getElementById(x),obs=()=>({name:$('name').value,lat:+$('lat').value,lon:+$('lon').value,model:$('model').value});
const viewpoints={{ viewpoints|tojson }};
function color(s){return s>=80?'#45e0a8':s>=65?'#7bd66a':s>=45?'#ffcc66':'#ff6b75'}
function fmtTime(t){return new Date(t).toLocaleTimeString('zh-CN',{hour:'2-digit',minute:'2-digit'})}
function outOfChina(lat,lon){return lon<72.004||lon>137.8347||lat<0.8293||lat>55.8271}
function transformLat(x,y){let r=-100+2*x+3*y+.2*y*y+.1*x*y+.2*Math.sqrt(Math.abs(x));r+=(20*Math.sin(6*x*Math.PI)+20*Math.sin(2*x*Math.PI))*2/3;r+=(20*Math.sin(y*Math.PI)+40*Math.sin(y/3*Math.PI))*2/3;r+=(160*Math.sin(y/12*Math.PI)+320*Math.sin(y*Math.PI/30))*2/3;return r}
function transformLon(x,y){let r=300+x+2*y+.1*x*x+.1*x*y+.1*Math.sqrt(Math.abs(x));r+=(20*Math.sin(6*x*Math.PI)+20*Math.sin(2*x*Math.PI))*2/3;r+=(20*Math.sin(x*Math.PI)+40*Math.sin(x/3*Math.PI))*2/3;r+=(150*Math.sin(x/12*Math.PI)+300*Math.sin(x/30*Math.PI))*2/3;return r}
function wgsToGcj(lat,lon){if(outOfChina(lat,lon))return[lat,lon];const a=6378245,ee=.00669342162296594323,dLat=transformLat(lon-105,lat-35),dLon=transformLon(lon-105,lat-35),rad=lat/180*Math.PI,magic=1-ee*Math.sin(rad)**2,sqrt=Math.sqrt(magic);return[lat+dLat*180/((a*(1-ee))/(magic*sqrt)*Math.PI),lon+dLon*180/(a/sqrt*Math.cos(rad)*Math.PI)]}
function gcjToWgs(lat,lon){if(outOfChina(lat,lon))return[lat,lon];const g=wgsToGcj(lat,lon);return[lat*2-g[0],lon*2-g[1]]}
function setViewpoint(i){const v=viewpoints[i];$('name').value=v.name;$('lat').value=v.lat.toFixed(6);$('lon').value=v.lon.toFixed(6);$('elev').value='读取中';map.closePopup();loadAll()}
viewpoints.forEach((v,i)=>{const p=wgsToGcj(v.lat,v.lon);L.circleMarker(p,{radius:6,color:'#38ef9d',weight:2,fillColor:'#0b513e',fillOpacity:.95}).bindPopup(`<b>${v.name}</b><br>${v.district} · ${v.target}<br><small>${v.precision}；出发前请以高德导航入口为准</small><br><button onclick="setViewpoint(${i})" style="margin-top:7px;padding:6px 8px">设为观测点并预报</button>`).addTo(viewpointLayer)});
let lastData=null;
function draw(d){lastData=d;layers.forEach(x=>map.removeLayer(x));layers=[];let o=d.observer,og=wgsToGcj(o.lat,o.lon),bounds=[og];$('elev').value=o.elev.toFixed(0);layers.push(L.circleMarker(og,{radius:8,color:'#43d9ff',fillOpacity:1}).addTo(map).bindPopup(`${o.name}<br>自动海拔 ${o.elev}m<br><small>原始坐标 WGS-84</small>`));d.mountains.forEach(m=>{let mg=wgsToGcj(m.lat,m.lon);bounds.push(mg);layers.push(L.polyline([og,mg],{color:'#4b8aa8',weight:2,dashArray:'6 7'}).addTo(map));layers.push(L.circleMarker(mg,{radius:7,color:'#ffcc66',fillColor:'#ffcc66',fillOpacity:.9}).addTo(map).bindPopup(`${m.name}<br>${m.elev}m · ${m.distance}km`))});map.fitBounds(bounds,{padding:[25,25]});
$('cards').innerHTML=d.mountains.map((m,mi)=>`<section class="panel mountain"><h3>${m.name}</h3><div class="meta">${m.distance} km · 方位 ${m.bearing}° · 峰顶仰角约 ${m.peak_angle}°</div><div class="days">${m.daily.map((x,di)=>{let b=x.morning;return `<div class="day"><strong>${x.date.slice(5)}</strong><div class="score" style="color:${color(b.score)}">${b.score}</div><div class="bar"><i style="width:${b.score}%"></i></div><div class="tabs"><span>晨 ${fmtTime(b.time)}</span><span>晚 ${x.evening.score}</span></div><div class="small">AOD550 ${b.aod==null?'暂无':b.aod.toFixed(2)}<br>PM2.5 ${b.pm2_5==null?'暂无':b.pm2_5.toFixed(1)} μg/m³<br>沙尘 ${b.dust==null?'暂无':b.dust.toFixed(1)} μg/m³<br>低云 ${b.low}% · 中云 ${b.mid}%<br>能见 ${b.visibility}km · RH ${b.rh}%<br>金山最高 ${x.gold.gold} 分<br>${b.reasons.join('、')}</div><div class="sim-actions"><button onclick="openSimulation(${mi},${di},'morning')">晨间形态</button><button onclick="openSimulation(${mi},${di},'evening')">傍晚形态</button></div></div>`}).join('')}</div></section>`).join('')}

function clamp01(x){return Math.max(0,Math.min(1,x))}
function humidityName(rh){return rh>=95?'近饱和':rh>=85?'高湿':rh>=70?'湿润':rh>=50?'适中':'干燥'}
function fogName(x){if(x.visibility<=1)return'浓雾风险';if(x.visibility<=5)return'雾/轻雾';if(x.visibility<=10||x.rh>=92)return'薄雾/霾感';return'无明显雾'}
function closeSimulation(){$('simModal').classList.remove('open')}
function openSimulation(mi,di,period){if(!lastData)return;const m=lastData.mountains[mi],x=m.daily[di][period];$('simTitle').textContent=`${m.name} · 路径云层剖面`;$('simSub').textContent=`${period==='morning'?'清晨':'傍晚'} ${new Date(x.time).toLocaleString('zh-CN')} · ${lastData.weather_model.name}`;$('simMetrics').innerHTML=`<div><b>${x.low}%</b>路径低云峰值</div><div><b>${x.mid}%</b>峰区中云</div><div><b>${x.high}%</b>路径高云均值</div><div><b>${x.rh}%</b>${humidityName(x.rh)}</div><div><b>${x.visibility} km</b>最低能见度</div><div><b>${x.aod==null?'—':x.aod.toFixed(2)}</b>AOD550</div>`;$('simNote').innerHTML=`${fogName(x)}；${x.reasons.join('、')}。<br>低云、中云、高云为模型分层云量；垂直高度带是近似范围，剖面用于判断视线遮挡，不代表实测云底云顶。`;$('simModal').classList.add('open');requestAnimationFrame(()=>drawPathProfile(x,m))}
function drawPathProfile(x,m){const canvas=$('simCanvas'),rect=canvas.getBoundingClientRect(),dpr=Math.min(2,devicePixelRatio||1);canvas.width=Math.round(rect.width*dpr);canvas.height=Math.round(rect.height*dpr);const c=canvas.getContext('2d');c.scale(dpr,dpr);const W=rect.width,H=rect.height,P=x.profile||[];if(P.length<2)return;const L=Math.max(48,W*.075),R=18,T=22,B=42,pw=W-L-R,ph=H-T-B,D=P[P.length-1].distance,maxTerrain=Math.max(...P.map(p=>p.terrain)),maxAlt=Math.max(16000,Math.ceil((maxTerrain+12000)/2000)*2000),Re=6371008.8*7/6,X=d=>L+d/D*pw,Y=z=>T+(maxAlt-z)/maxAlt*ph,bulge=d=>d*1000*(D-d)*1000/(2*Re),effective=p=>p.terrain+bulge(p.distance);c.fillStyle='#03070c';c.fillRect(0,0,W,H);c.fillStyle='#06111a';c.fillRect(L,T,pw,ph);
// 高度与距离网格
c.font=`${Math.max(10,H*.027)}px system-ui`;c.lineWidth=1;c.textAlign='right';for(let z=0;z<=maxAlt;z+=2000){const y=Y(z);c.strokeStyle='rgba(130,165,185,.22)';c.beginPath();c.moveTo(L,y);c.lineTo(W-R,y);c.stroke();c.fillStyle='#a9bac6';c.fillText(z+'m',L-6,y+4)}c.textAlign='center';for(let i=0;i<=4;i++){const d=D*i/4,xx=X(d);c.strokeStyle='rgba(130,165,185,.14)';c.beginPath();c.moveTo(xx,T);c.lineTo(xx,H-B);c.stroke();c.fillStyle='#a9bac6';c.fillText(Math.round(d)+'km',xx,H-B+18)}
// 各路径格点的气溶胶薄层、云层与雾
const bands=[['high',7000,12000,'214,204,246'],['mid',3000,7000,'186,194,207'],['low',200,3000,'116,151,180']];let blocked=[];for(let i=0;i<P.length;i++){const p=P[i],d0=i===0?0:(P[i-1].distance+p.distance)/2,d1=i===P.length-1?D:(p.distance+P[i+1].distance)/2,x0=X(d0),x1=X(d1),ground=effective(p),aod=clamp01((p.aod||0)/.6);if(aod>.03){c.fillStyle=`rgba(191,166,127,${.04+aod*.18})`;c.fillRect(x0,T,x1-x0,ph)}for(const [key,lo,hi,rgb] of bands){const cover=p[key]||0,base=ground+lo,top=Math.min(maxAlt,ground+hi);if(cover>2&&top>base){c.fillStyle=`rgba(${rgb},${.06+cover/135})`;c.fillRect(x0,Y(top),x1-x0,Y(base)-Y(top));c.strokeStyle=`rgba(${rgb},${.18+cover/150})`;c.strokeRect(x0,Y(top),x1-x0,Y(base)-Y(top));const los=P[0].terrain+(P[P.length-1].terrain-P[0].terrain)*(p.distance/D);if(cover>=45&&los>=base&&los<=top)blocked.push({d:p.distance,z:los,key,cover})}}const fog=clamp01((p.rh-86)/14+(8-p.visibility)/10);if(fog>0){const fh=250+fog*900;c.fillStyle=`rgba(214,235,231,${.14+fog*.6})`;c.fillRect(x0,Y(ground+fh),x1-x0,Y(ground)-Y(ground+fh))}}
// 地形剖面（含7/6地球半径折射修正）
c.beginPath();c.moveTo(X(P[0].distance),Y(effective(P[0])));P.forEach(p=>c.lineTo(X(p.distance),Y(effective(p))));c.lineTo(X(D),Y(0));c.lineTo(X(0),Y(0));c.closePath();let tg=c.createLinearGradient(0,Y(maxTerrain+1000),0,Y(0));tg.addColorStop(0,'#6f8567');tg.addColorStop(1,'#26352e');c.fillStyle=tg;c.fill();c.beginPath();P.forEach((p,i)=>i?c.lineTo(X(p.distance),Y(effective(p))):c.moveTo(X(p.distance),Y(effective(p))));c.strokeStyle='#72d3a1';c.lineWidth=2;c.stroke();
// 观测视线：观测点直达峰顶
const y0=Y(P[0].terrain),y1=Y(P[P.length-1].terrain);c.setLineDash([8,6]);c.strokeStyle='#ffe95a';c.lineWidth=2.4;c.beginPath();c.moveTo(X(0),y0);c.lineTo(X(D),y1);c.stroke();c.setLineDash([]);c.fillStyle='#45e0c0';c.beginPath();c.arc(X(0),y0,6,0,Math.PI*2);c.fill();c.fillStyle='#ffe95a';c.beginPath();c.moveTo(X(D),y1-9);c.lineTo(X(D)-7,y1+6);c.lineTo(X(D)+7,y1+6);c.closePath();c.fill();
// 云层与视线交点
blocked.forEach(b=>{c.fillStyle='#ff5d6c';c.beginPath();c.arc(X(b.d),Y(b.z),4,0,Math.PI*2);c.fill()});c.textAlign='left';c.fillStyle='#d7e5ed';c.fillText('观测点',X(0)+8,y0-8);c.textAlign='right';c.fillStyle='#ffe95a';c.fillText(`${m.name} ${m.elev}m`,X(D)-5,y1-10);c.textAlign='left';c.fillStyle='#9bb3c3';c.fillText('有效地形：已考虑地球曲率与标准折射',L,T+13);
// 图例
const legend=[['#74a0c0','低云 0.2–3km AGL'],['#bac2cf','中云 3–7km AGL'],['#d6ccf6','高云 7–12km AGL'],['#d6ebe7','雾/高湿层'],['#ff5d6c','视线潜在遮挡']];let lx=L,ly=H-8;c.font=`${Math.max(9,H*.022)}px system-ui`;c.textAlign='left';legend.forEach(([co,tx])=>{c.fillStyle=co;c.fillRect(lx,ly-9,10,7);c.fillStyle='#b9cbd6';c.fillText(tx,lx+13,ly);lx+=c.measureText(tx).width+27;if(lx>W-145){lx=L;ly-=14}});if(blocked.length)$('simNote').innerHTML+=`<br><span style="color:#ff7c86">检测到 ${blocked.length} 个路径格点的云层高度带与观测视线相交，请重点关注红点位置。</span>`;else $('simNote').innerHTML+=`<br><span style="color:#62e6b3">当前近似云层高度带未与峰顶观测视线直接相交。</span>`}
let forecastRequest=0;
async function loadAll(){const id=++forecastRequest,chosen=$('model').value;localStorage.setItem('snowWeatherModel',chosen);document.body.classList.add('loading');$('modelStatus').textContent='天气模型：正在获取 '+$('model').selectedOptions[0].text+'…';$('aerosol').textContent='气溶胶：正在获取 Open-Meteo AOD550…';try{let q=new URLSearchParams(obs()),r=await fetch('/api/forecast?'+q),j=await r.json();if(id!==forecastRequest)return;if(!j.ok)throw Error(j.error);draw(j.data);$('modelStatus').textContent='天气模型：'+j.data.weather_model.name+' · '+j.data.weather_model.detail;$('modelStatus').style.borderColor='#43d9ff';$('aerosol').textContent='气溶胶：'+j.data.aerosol.message;$('aerosol').style.borderColor='#45e0a8'}catch(e){if(id!==forecastRequest)return;$('modelStatus').textContent='数据源获取失败：'+e.message;$('modelStatus').style.borderColor='#ff6b75';$('aerosol').textContent='获取失败：'+e.message;$('aerosol').style.borderColor='#ff6b75';alert('预报失败：'+e.message)}finally{if(id===forecastRequest)document.body.classList.remove('loading')}}
let elevTimer,elevRequest=0;
async function loadElevationGrid(){if(!map.hasLayer(elevationLayer))return;const id=++elevRequest,b=map.getBounds(),sw=gcjToWgs(b.getSouth(),b.getWest()),ne=gcjToWgs(b.getNorth(),b.getEast()),q=new URLSearchParams({south:sw[0],west:sw[1],north:ne[0],east:ne[1],n:12});try{const j=await fetch('/api/elevation-grid?'+q).then(r=>r.json());if(id!==elevRequest||!j.ok)return;elevationLayer.clearLayers();const colors={low:'#38c977',mid:'#ffb52e',high:'#b968e8',unknown:'#777'};j.cells.forEach(c=>{const a=wgsToGcj(c.south,c.west),z=wgsToGcj(c.north,c.east);L.rectangle([a,z],{pane:'terrainPane',stroke:false,fillColor:colors[c.band],fillOpacity:.24,interactive:true}).bindTooltip(`${c.elevation==null?'未知':Math.round(c.elevation)+' m'} · ${c.band==='low'?'低海拔':c.band==='mid'?'中海拔':c.band==='high'?'高海拔':'未知'}`).addTo(elevationLayer)})}catch(e){console.warn('海拔分区获取失败',e)}}
map.on('moveend',()=>{clearTimeout(elevTimer);elevTimer=setTimeout(loadElevationGrid,500)});map.on('overlayadd',e=>{if(e.layer===elevationLayer)loadElevationGrid()});
function locate(){navigator.geolocation?navigator.geolocation.getCurrentPosition(p=>{$('lat').value=p.coords.latitude.toFixed(6);$('lon').value=p.coords.longitude.toFixed(6);$('elev').value='读取中';loadAll()},e=>alert(e.message),{enableHighAccuracy:true}):alert('浏览器不支持定位')}
const savedModel=localStorage.getItem('snowWeatherModel');if(savedModel&&[...$('model').options].some(x=>x.value===savedModel))$('model').value=savedModel;loadAll();
</script></body></html>'''


if __name__ == "__main__":
    print("成都看雪山：http://127.0.0.1:5000")
    APP.run(host="0.0.0.0", port=int(os.getenv("PORT", "5000")), debug=False, threaded=True)
