#!/usr/bin/env python3
"""ECOS backend: REST API + static frontend. Python 3.9+ standard library only (SQLite storage).

Run:  python import_excel.py   (once, builds ecos.db)   then   python server.py
Open: http://localhost:8000
"""
import json
import mimetypes
import os
import re
import sqlite3
import sys
from datetime import date, datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import parse_qs, unquote, urlparse

import summarizer
from schema import TABLES

BASE = os.path.dirname(os.path.abspath(__file__))


def load_env_file(path):
    """Read KEY=VALUE lines from backend/.env (e.g. ANTHROPIC_API_KEY). Real environment variables take priority."""
    try:
        with open(path, encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if line and not line.startswith("#") and "=" in line:
                    k, v = line.split("=", 1)
                    os.environ.setdefault(k.strip(), v.strip().strip("\"'"))
    except OSError:
        pass


load_env_file(os.path.join(BASE, ".env"))
DB = os.environ.get("ECOS_DB", os.path.join(BASE, "ecos.db"))
FRONT = os.path.abspath(os.environ.get("ECOS_FRONTEND", os.path.join(BASE, "..", "frontend")))
MAX_BODY = 1_000_000
PRIO = {"Critical": 0, "High": 1, "Medium": 2, "Low": 3}


def today():
    return os.environ.get("ECOS_TODAY") or date.today().isoformat()


class ApiError(Exception):
    def __init__(self, status, code, message, details=None):
        super().__init__(message)
        self.status, self.code, self.message, self.details = status, code, message, details


def connect():
    c = sqlite3.connect(DB, timeout=10, isolation_level=None)
    c.row_factory = sqlite3.Row
    c.execute("PRAGMA foreign_keys=ON")
    return c


def q(c, sql, args=()):
    return [dict(r) for r in c.execute(sql, args)]


def rnd(x):
    return int(x + 0.5)


# ---------------------------------------------------------------- validation
def validate(c, route, data, partial, current=None):
    spec = TABLES[route]
    if not isinstance(data, dict):
        raise ApiError(400, "bad_request", "Request body must be a JSON object")
    errs = {}
    unknown = [k for k in data if k not in spec["cols"]]
    if unknown:
        raise ApiError(400, "unknown_fields", "Unknown fields: " + ", ".join(unknown), {"allowed": list(spec["cols"])})
    out = {}
    for col, val in data.items():
        typ = spec["cols"][col]
        if col == spec["pk"] and partial:
            if val != (current or {}).get(col):
                errs[col] = "Primary key cannot be changed"
            continue
        try:
            if typ == "int":
                if isinstance(val, bool) or (isinstance(val, float) and val != int(val)) or isinstance(val, (list, dict)):
                    raise ValueError
                val = int(val)
            elif typ == "real":
                if isinstance(val, bool) or isinstance(val, (list, dict)):
                    raise ValueError
                val = round(float(val), 4)
            else:
                if not isinstance(val, str):
                    raise ValueError
                val = val.strip()
                if not val:
                    errs[col] = "Must not be empty"
                    continue
        except (ValueError, TypeError):
            errs[col] = f"Must be a valid {'whole number' if typ == 'int' else 'number' if typ == 'real' else 'text value'}"
            continue
        if col in spec["enums"] and val not in spec["enums"][col]:
            errs[col] = "Must be one of: " + ", ".join(spec["enums"][col])
        if col in spec["ranges"]:
            lo, hi = spec["ranges"][col]
            if not lo <= val <= hi:
                errs[col] = f"Must be between {lo} and {hi}"
        kind = spec["kinds"].get(col)
        try:
            if kind == "date":
                date.fromisoformat(val)
            elif kind == "datetime":
                datetime.strptime(val, "%Y-%m-%d %H:%M")
        except ValueError:
            errs[col] = "Must match " + ("YYYY-MM-DD" if kind == "date" else "YYYY-MM-DD HH:MM")
        if col in spec["fks"] and col not in errs:
            r2, pk2 = spec["fks"][col]
            if not c.execute(f'SELECT 1 FROM "{TABLES[r2]["table"]}" WHERE "{pk2}"=?', (val,)).fetchone():
                errs[col] = f"{val} does not exist in {r2}"
        out[col] = val
    if not partial:
        for col in spec["required"]:
            if col not in out and col not in errs:
                errs[col] = "Required"
    merged = {**(current or {}), **out}
    if route == "tasks" and "Dependency_Task_ID" in out:
        d = out["Dependency_Task_ID"]
        if d != "None" and not c.execute('SELECT 1 FROM Tasks WHERE Task_ID=?', (d,)).fetchone():
            errs["Dependency_Task_ID"] = 'Must be "None" or an existing task id'
        elif d != "None" and d == merged.get("Task_ID"):
            errs["Dependency_Task_ID"] = "A task cannot depend on itself"
    if route == "tasks" and not partial and "Dependency_Task_ID" not in out:
        out["Dependency_Task_ID"] = "None"
    if route == "meetings" and "Participants" in out:
        ids = [p.strip() for p in out["Participants"].split(",") if p.strip()]
        miss = [i for i in ids if not c.execute("SELECT 1 FROM Employees WHERE Employee_ID=?", (i,)).fetchone()]
        if not ids or miss:
            errs["Participants"] = "Comma-separated existing employee ids required" + (f"; unknown: {', '.join(miss)}" if miss else "")
        else:
            out["Participants"] = ", ".join(ids)
    if route == "employees" and "Skills" in out:
        out["Skills"] = ", ".join(s.strip() for s in out["Skills"].split(",") if s.strip())
    if route == "dependencies" and merged.get("Predecessor_Task_ID") and merged.get("Predecessor_Task_ID") == merged.get("Successor_Task_ID"):
        errs["Successor_Task_ID"] = "Predecessor and successor must differ"
    if errs:
        raise ApiError(422, "validation_failed", "Validation failed", errs)
    return out


def next_id(c, spec):
    ids = [r[0] for r in c.execute(f'SELECT "{spec["pk"]}" FROM "{spec["table"]}"')]
    ms = [m for m in (re.match(r"^([A-Za-z]+)(\d+)$", i) for i in ids) if m]
    if not ms:
        raise ApiError(400, "id_required", f'Provide {spec["pk"]} explicitly')
    prefix = ms[0].group(1)
    width = max(len(m.group(2)) for m in ms)
    n = max(int(m.group(2)) for m in ms if m.group(1) == prefix) + 1
    return f"{prefix}{n:0{width}d}"


# ---------------------------------------------------------------- domain model: CSP + heuristic search
class Ctx:
    """Snapshot of the data needed to evaluate constraints."""

    def __init__(self, c):
        self.today = today()
        self.emps = {}
        for r in q(c, "SELECT * FROM Employees"):
            self.emps[r["Employee_ID"]] = dict(id=r["Employee_ID"], name=r["Employee_Name"], dept=r["Department"], role=r["Role"],
                                              skills=[s.strip() for s in r["Skills"].split(",") if s.strip()], exp=r["Experience_Years"],
                                              hours=r["Availability_Hours_Per_Day"], workload=r["Current_Workload_Percent"], perf=r["Performance_Score"],
                                              team=r["Team_ID"], mode=r["Work_Mode"])
        self.tasks = {r["Task_ID"]: r for r in q(c, "SELECT * FROM Tasks")}
        self.assigned, self.active = {}, {}
        for a in q(c, "SELECT Employee_ID,Task_ID FROM Assignments WHERE Assignment_Status='Assigned'"):
            self.assigned.setdefault(a["Task_ID"], []).append(a["Employee_ID"])
            self.active[a["Employee_ID"]] = self.active.get(a["Employee_ID"], 0) + 1
        self.busy = {}
        for r in q(c, "SELECT a.Employee_ID e FROM Assignments a JOIN Tasks t ON t.Task_ID=a.Task_ID WHERE a.Assignment_Status='Assigned' AND t.Status!='Completed'"):
            self.busy[r["e"]] = self.busy.get(r["e"], 0) + 1
        self.res = {}
        for r in q(c, "SELECT Project_ID,Availability_Status FROM Resources"):
            self.res.setdefault(r["Project_ID"], []).append(r["Availability_Status"])

    def assignees(self, task_id):
        return [self.emps[i] for i in self.assigned.get(task_id, []) if i in self.emps]


def check(cn, emps, task, ctx):
    """True when constraint cn holds for the given employees working on task."""
    t, v = cn["Constraint_Type"], str(cn["Constraint_Value"])
    m = re.search(r"\d+", v)
    n = float(m.group()) if m else 0
    if t == "Required Skill":
        return all(v in e["skills"] for e in emps)
    if t == "Minimum Experience":
        return all(e["exp"] >= n for e in emps)
    if t == "Maximum Workload":
        return all(e["workload"] <= n for e in emps)
    if t == "Availability":
        return all(e["hours"] >= 5 for e in emps)
    if t == "Working Hours":
        return all(e["hours"] <= n for e in emps)
    if t == "Department Match":
        return all(e["dept"] == task["Department"] for e in emps)
    if t == "Deadline":
        return task["Status"] == "Completed" or task["Deadline"] >= ctx.today
    if t == "Task Dependency":
        d = task["Dependency_Task_ID"]
        return d == "None" or d not in ctx.tasks or ctx.tasks[d]["Status"] == "Completed"
    if t == "Team Capacity":
        return all(ctx.active.get(e["id"], 0) <= n for e in emps)
    if t == "Resource Availability":
        rs = ctx.res.get(task["Project_ID"], [])
        return not (rs and all(s == "Unavailable" for s in rs))
    return True


def with_status(c, rows, ctx=None):
    ctx = ctx or Ctx(c)
    for r in rows:
        t = ctx.tasks.get(r["Task_ID"])
        r["Status"] = "Satisfied" if not t or check(r, ctx.assignees(r["Task_ID"]), t, ctx) else "Violated"
    return rows


def score(e, task):
    skills = [s.strip() for s in task["Required_Skills"].split(",") if s.strip()]
    match = rnd(100 * len([k for k in skills if k in e["skills"]]) / max(1, len(skills)))
    avail = rnd(e["hours"] / 9 * 100)
    wl = max(0, 100 - e["workload"] * 2)
    return match, avail, wl, rnd(.4 * match + .25 * avail + .2 * wl + .15 * e["perf"])


def hard_violations(c, ctx, task, emp):
    cons = q(c, "SELECT * FROM Constraints WHERE Task_ID=? AND Constraint_Level='Hard'", (task["Task_ID"],))
    return [f'{k["Constraint_Type"]}: {k["Constraint_Value"]}' for k in cons if not check(k, [emp], task, ctx)]


def candidates(c, ctx, task):
    """Generate -> score -> constraint-check every employee not already on the task, best first."""
    cons = q(c, "SELECT * FROM Constraints WHERE Task_ID=? AND Constraint_Level='Hard'", (task["Task_ID"],))
    on_task = set(ctx.assigned.get(task["Task_ID"], []))
    cands = []
    for e in ctx.emps.values():
        if e["id"] in on_task:
            continue
        match, avail, wl, sc = score(e, task)
        viol = [f'{k["Constraint_Type"]}: {k["Constraint_Value"]}' for k in cons if not check(k, [e], task, ctx)]
        cands.append(dict(employee_id=e["id"], name=e["name"], department=e["dept"], skill_match=match, availability_hours=e["hours"],
                          availability_score=avail, open_tasks=ctx.busy.get(e["id"], 0),
                          available=e["hours"] >= 5 and ctx.busy.get(e["id"], 0) == 0, workload=e["workload"], performance=e["perf"], score=sc,
                          passes_hard_constraints=not viol, violations=viol))
    cands.sort(key=lambda x: (not x["passes_hard_constraints"], -x["score"]))
    return cands


def optimize(c, task_id, top=None, n_options=3):
    ctx = Ctx(c)
    task = ctx.tasks.get(task_id)
    if not task:
        raise ApiError(404, "not_found", f"Task {task_id} not found")
    on_task = ctx.assigned.get(task_id, [])
    remaining = max(0, task["Required_Employees"] - len(on_task))
    cands = candidates(c, ctx, task)
    # options: the best free employees (>=5 h/day, no open task) that pass every hard constraint; one is chosen per assignment
    options = [x for x in cands if x["passes_hard_constraints"] and x["available"]][:n_options] if remaining else []
    rec = [options[0]["employee_id"]] if options else []
    for i, x in enumerate(cands, 1):
        x["rank"], x["recommended"] = i, x["employee_id"] in rec
    for i, x in enumerate(options, 1):
        x["option"] = i
    top = max(6, remaining + 3) if top is None else top
    return dict(task=task, strategy="Weighted heuristic (skill match 40%, availability 25%, workload headroom 20%, performance 15%) then CSP hard-constraint filter",
                evaluated=len(cands), feasible=sum(1 for x in cands if x["passes_hard_constraints"]),
                required=task["Required_Employees"], assigned=[{"employee_id": i, "name": ctx.emps[i]["name"]} for i in on_task if i in ctx.emps],
                remaining=remaining, free=sum(1 for x in cands if x["available"]), options=options, recommended_ids=rec,
                candidates=cands[:top], best=cands[0] if cands else None)


def apply_assignment(c, data):
    """Assign employees to a task. Use employee_ids=[...] / employee_id, or auto=true for the single best free employee.
    An employee is free with >=5 h/day and no other open (non-completed) Assigned task; force=true overrides the checks."""
    if not isinstance(data, dict) or not data.get("task_id"):
        raise ApiError(422, "validation_failed", "task_id is required")
    ids = data.get("employee_ids") if isinstance(data.get("employee_ids"), list) else ([data["employee_id"]] if data.get("employee_id") else [])
    if not ids and not data.get("auto"):
        raise ApiError(422, "validation_failed", "Provide employee_ids, employee_id or auto=true")
    c.execute("BEGIN IMMEDIATE")
    try:
        ctx = Ctx(c)
        task = ctx.tasks.get(data["task_id"])
        if not task:
            raise ApiError(404, "not_found", f'Task {data["task_id"]} not found')
        if task["Status"] == "Completed":
            raise ApiError(409, "task_completed", "Task is already completed")
        on_task = set(ctx.assigned.get(task["Task_ID"], []))
        remaining = max(0, task["Required_Employees"] - len(on_task))
        if data.get("auto"):
            ids = [x["employee_id"] for x in candidates(c, ctx, task) if x["passes_hard_constraints"] and x["available"]][:1] if remaining else []
            if not ids:
                raise ApiError(409, "no_candidates", "No free employee satisfies the hard constraints, or the task is already fully staffed")
        ids = list(dict.fromkeys(str(i) for i in ids))
        emps = []
        for i in ids:
            if i not in ctx.emps:
                raise ApiError(404, "not_found", f"Employee {i} not found")
            if i in on_task:
                raise ApiError(409, "duplicate_assignment", f'{ctx.emps[i]["name"]} is already assigned to {task["Task_ID"]}')
            emps.append(ctx.emps[i])
        viol = [f'{e["name"]} - {v}' for e in emps for v in hard_violations(c, ctx, task, e)]
        if viol and not data.get("force"):
            raise ApiError(409, "constraint_violation", "Assignment violates hard constraints", {"violations": viol})
        busy = [f'{e["name"]} {"has " + str(ctx.busy[e["id"]]) + " open task(s)" if ctx.busy.get(e["id"]) else "has under 5 free hours a day"}' for e in emps
                if ctx.busy.get(e["id"], 0) or e["hours"] < 5]
        if busy and not data.get("force"):
            raise ApiError(409, "employee_busy", "Employee is not free: one person takes one open task at a time", {"busy": busy})
        if len(emps) > remaining and not data.get("force"):
            raise ApiError(409, "too_many" if remaining else "task_full",
                           f'Task needs {remaining} more employee(s) but {len(emps)} were selected' if remaining else "Task already has all the employees it needs")
        spec, created = TABLES["assignments"], []
        for e in emps:
            match, avail, wl, sc = score(e, task)
            row = {"Assignment_ID": next_id(c, spec), "Employee_ID": e["id"], "Task_ID": task["Task_ID"], "Skill_Match_Score": match / 100,
                   "Availability_Score": avail / 100, "Workload_Factor": round(wl / 100, 2), "Assignment_Score": float(sc), "Assignment_Status": "Assigned"}
            c.execute(f'INSERT INTO Assignments ({",".join(row)}) VALUES ({",".join("?" * len(row))})', list(row.values()))
            created.append(row)
        if task["Status"] == "Pending":
            c.execute("UPDATE Tasks SET Status='In Progress' WHERE Task_ID=?", (task["Task_ID"],))
        c.execute("COMMIT")
    except BaseException:
        if c.in_transaction:
            c.execute("ROLLBACK")
        raise
    return dict(assignments=created, assignment=created[0], task=q(c, "SELECT * FROM Tasks WHERE Task_ID=?", (task["Task_ID"],))[0],
                still_needed=max(0, remaining - len(created)), warnings=viol)


# ---------------------------------------------------------------- reports
def dashboard(c):
    ctx = Ctx(c)
    tasks = list(ctx.tasks.values())
    td = date.fromisoformat(ctx.today)
    open_ = [t for t in tasks if t["Status"] != "Completed"]
    live = q(c, "SELECT Assignment_Score s FROM Assignments WHERE Assignment_Status!='Rejected'")
    cons = with_status(c, q(c, "SELECT * FROM Constraints"), ctx)
    bad = {k["Task_ID"] for k in cons if k["Status"] == "Violated" and k["Constraint_Level"] == "Hard"}
    emps = list(ctx.emps.values())
    cnt = lambda s: sum(1 for t in tasks if t["Status"] == s)
    return dict(
        total_employees=len(emps), total_projects=c.execute("SELECT COUNT(*) FROM Projects").fetchone()[0],
        active_projects=c.execute("SELECT COUNT(*) FROM Projects WHERE Status='Active'").fetchone()[0],
        pending_tasks=cnt("Pending"), in_progress_tasks=cnt("In Progress"), completed_tasks=cnt("Completed"), blocked_tasks=cnt("Blocked"),
        high_priority_open=sum(1 for t in open_ if t["Priority"] in ("Critical", "High")),
        available_resources=c.execute("SELECT COUNT(*) FROM Resources WHERE Availability_Status='Available'").fetchone()[0],
        avg_workload=rnd(sum(e["workload"] for e in emps) / len(emps)),
        optimization_score=rnd(sum(r["s"] for r in live) / len(live)) if live else 0,
        insights=[
            f'{sum(1 for t in open_ if 0 <= (date.fromisoformat(t["Deadline"]) - td).days <= 14)} open tasks are due within 14 days.',
            f'{sum(1 for e in emps if e["workload"] < 15)} employees are below 15% workload.',
            f'{sum(1 for t in open_ if t["Priority"] in ("Critical", "High") and t["Task_ID"] in bad)} high-priority tasks have unresolved hard-constraint violations.',
            f'{c.execute("SELECT COUNT(*) FROM Assignments WHERE Assignment_Status=? AND Assignment_Score>=85", ("Suggested",)).fetchone()[0]} suggested assignments score 85 or higher and are ready to apply.'])


def analytics(c):
    ctx = Ctx(c)
    pct = lambda n, d: round(100 * n / d, 1) if d else 0
    tasks = list(ctx.tasks.values())
    cons = with_status(c, q(c, "SELECT * FROM Constraints"), ctx)
    asg = q(c, "SELECT Assignment_Status s, COUNT(*) n FROM Assignments GROUP BY 1")
    live = q(c, "SELECT Assignment_Score s FROM Assignments WHERE Assignment_Status!='Rejected'")
    done = pct(sum(1 for t in tasks if t["Status"] == "Completed"), len(tasks))
    viol = pct(sum(1 for k in cons if k["Status"] == "Violated"), len(cons))
    by = lambda rows, k: {x[k]: x["n"] for x in rows}
    ok = sum(a["n"] for a in asg if a["s"] in ("Assigned", "Completed"))
    return dict(
        task_completion_rate=done, constraint_violation_rate=viol, assignment_success_rate=pct(ok, sum(a["n"] for a in asg)),
        collaboration_optimization_score=rnd(.5 * sum(r["s"] for r in live) / len(live) + .3 * (100 - viol) + .2 * done),
        tasks_by_status=by(q(c, "SELECT Status k,COUNT(*) n FROM Tasks GROUP BY 1"), "k"), tasks_by_priority=by(q(c, "SELECT Priority k,COUNT(*) n FROM Tasks GROUP BY 1"), "k"),
        assignments_by_status=by([{"k": a["s"], "n": a["n"]} for a in asg], "k"),
        department_productivity={r["k"]: pct(r["d"], r["n"]) for r in q(c, "SELECT Department k,COUNT(*) n,SUM(Status='Completed') d FROM Tasks GROUP BY 1")},
        workload_by_department={r["k"]: round(r["a"], 1) for r in q(c, "SELECT Department k,AVG(Current_Workload_Percent) a FROM Employees GROUP BY 1")},
        resource_utilization_by_type={r["k"]: round(r["a"], 1) for r in q(c, "SELECT Resource_Type k,AVG(Utilization_Percent) a FROM Resources GROUP BY 1")},
        project_progress={r["Project_ID"]: pct(r["d"] or 0, r["n"]) for r in q(c, "SELECT Project_ID,COUNT(*) n,SUM(Status='Completed') d FROM Tasks GROUP BY 1")},
        violations_by_type={t: pct(sum(1 for k in cons if k["Constraint_Type"] == t and k["Status"] == "Violated"), sum(1 for k in cons if k["Constraint_Type"] == t))
                            for t in sorted({k["Constraint_Type"] for k in cons})})


def recommendations(c):
    ctx = Ctx(c)
    td = date.fromisoformat(ctx.today)
    emps, tasks = ctx.emps, ctx.tasks
    out = {"task": [], "workforce": [], "deadline_risk": [], "resource": []}
    for a in q(c, "SELECT * FROM Assignments WHERE Assignment_Status='Suggested' ORDER BY Assignment_Score DESC LIMIT 4"):
        out["task"].append(dict(id=a["Assignment_ID"], recommendation=f'Assign {emps[a["Employee_ID"]]["name"]} to {a["Task_ID"]}',
                                reason=f'Skill match {rnd(a["Skill_Match_Score"] * 100)}%, availability {rnd(a["Availability_Score"] * 100)}%.',
                                impact=f'Better fit for {tasks[a["Task_ID"]]["Task_Name"]}', confidence=rnd(a["Assignment_Score"]), related=a["Task_ID"],
                                action=dict(method="PUT", path=f'/api/assignments/{a["Assignment_ID"]}', body={"Assignment_Status": "Assigned"})))
    for e in sorted(emps.values(), key=lambda e: e["workload"])[:3]:
        out["workforce"].append(dict(id=e["id"], recommendation=f'{e["name"]} ({e["id"]}) has spare capacity', reason=f'Workload is {e["workload"]}% with {e["hours"]} h/day available.',
                                     impact="Can absorb more tasks", confidence=min(99, e["perf"] + 5), related=e["dept"]))
    risky = [t for t in tasks.values() if t["Status"] in ("Blocked", "Pending") and 0 <= (date.fromisoformat(t["Deadline"]) - td).days <= 14]
    for t in sorted(risky, key=lambda t: (PRIO[t["Priority"]], t["Deadline"]))[:4]:
        d = (date.fromisoformat(t["Deadline"]) - td).days
        out["deadline_risk"].append(dict(id=t["Task_ID"], recommendation=f'{t["Task_ID"]} ({t["Task_Name"]}) may miss its deadline', reason=f'{t["Status"]} with {d} days left.',
                                         impact=f'{t["Priority"]} priority in {t["Project_ID"]}', confidence=70 + max(0, 14 - d), related=t["Project_ID"]))
    for r in q(c, "SELECT * FROM Resources ORDER BY Utilization_Percent DESC LIMIT 3"):
        out["resource"].append(dict(id=r["Resource_ID"], recommendation=f'{r["Resource_ID"]} ({r["Resource_Name"]}) is near maximum utilization', reason=f'Utilization is {r["Utilization_Percent"]}% on {r["Project_ID"]}.',
                                    impact="Rebalancing may reduce delay", confidence=r["Utilization_Percent"], related=r["Project_ID"]))
    return out


def employee_report(c, eid):
    """Everything ECOS knows about one employee, with rule-based insights (no language model)."""
    ctx = Ctx(c)
    e = ctx.emps.get(eid)
    if not e:
        raise ApiError(404, "not_found", f"Employee {eid} not found")
    td = date.fromisoformat(ctx.today)
    pn = {r["Project_ID"]: r["Project_Name"] for r in q(c, "SELECT Project_ID,Project_Name FROM Projects")}
    cons = with_status(c, q(c, "SELECT * FROM Constraints"), ctx)
    cons_by_task = {}
    for k in cons:
        cons_by_task.setdefault(k["Task_ID"], []).append(k)

    # --- assignments and tasks
    rows = q(c, "SELECT * FROM Assignments WHERE Employee_ID=? ORDER BY Assignment_ID", (eid,))
    by_status = {}
    for a in rows:
        by_status[a["Assignment_Status"]] = by_status.get(a["Assignment_Status"], 0) + 1

    def task_row(a):
        t = ctx.tasks[a["Task_ID"]]
        left = (date.fromisoformat(t["Deadline"]) - td).days
        ks = cons_by_task.get(t["Task_ID"], [])
        return dict(assignment_id=a["Assignment_ID"], assignment_status=a["Assignment_Status"], assignment_score=a["Assignment_Score"],
                    task_id=t["Task_ID"], task=t["Task_Name"], project_id=t["Project_ID"], project=pn.get(t["Project_ID"], t["Project_ID"]),
                    priority=t["Priority"], status=t["Status"], deadline=t["Deadline"], days_left=left, estimated_hours=t["Estimated_Hours"],
                    required_skills=t["Required_Skills"], skills_matched=[k.strip() for k in t["Required_Skills"].split(",") if k.strip() in e["skills"]],
                    constraints_violated=[f'{k["Constraint_Type"]}: {k["Constraint_Value"]}' for k in ks if k["Status"] == "Violated" and k["Constraint_Level"] == "Hard"])
    mine = [task_row(a) for a in rows if a["Task_ID"] in ctx.tasks]
    current = [t for t in mine if t["assignment_status"] == "Assigned" and t["status"] != "Completed"]
    current.sort(key=lambda t: (PRIO[t["priority"]], t["deadline"]))
    suggested = sorted((t for t in mine if t["assignment_status"] == "Suggested"), key=lambda t: -t["assignment_score"])
    finished = [t for t in mine if t["assignment_status"] == "Completed" or (t["assignment_status"] == "Assigned" and t["status"] == "Completed")]
    scored = [a["Assignment_Score"] for a in rows if a["Assignment_Status"] != "Rejected"]
    overdue = [t for t in current if t["days_left"] < 0]
    due_soon = [t for t in current if 0 <= t["days_left"] <= 14]
    at_risk = [t for t in current if t["constraints_violated"]]

    # --- comparison with the department and the organisation
    peers = [x for x in ctx.emps.values() if x["dept"] == e["dept"]]
    avg = lambda xs: round(sum(xs) / len(xs), 1) if xs else 0
    everyone = list(ctx.emps.values())
    perf_rank = 1 + sum(1 for x in everyone if x["perf"] > e["perf"])
    compare = dict(department=e["dept"], department_size=len(peers),
                   workload=dict(employee=e["workload"], department=avg([x["workload"] for x in peers]), organization=avg([x["workload"] for x in everyone])),
                   performance=dict(employee=e["perf"], department=avg([x["perf"] for x in peers]), organization=avg([x["perf"] for x in everyone])),
                   experience=dict(employee=e["exp"], department=avg([x["exp"] for x in peers]), organization=avg([x["exp"] for x in everyone])),
                   availability=dict(employee=e["hours"], department=avg([x["hours"] for x in peers]), organization=avg([x["hours"] for x in everyone])),
                   performance_rank=perf_rank, organization_size=len(everyone))

    # --- communication and meetings
    msgs = q(c, "SELECT * FROM Communications WHERE Sender_ID=? OR Receiver_ID=? ORDER BY Timestamp DESC", (eid, eid))
    names = {i: x["name"] for i, x in ctx.emps.items()}
    sent = [m for m in msgs if m["Sender_ID"] == eid]
    got = [m for m in msgs if m["Receiver_ID"] == eid]
    waiting = [m for m in got if m["Action_Required"] == "Yes"]
    recent = [dict(message_id=m["Message_ID"], direction="Sent" if m["Sender_ID"] == eid else "Received",
                   other=names.get(m["Receiver_ID"] if m["Sender_ID"] == eid else m["Sender_ID"], ""), type=m["Communication_Type"],
                   project=pn.get(m["Project_ID"], m["Project_ID"]), priority=m["Priority"], action_required=m["Action_Required"],
                   timestamp=m["Timestamp"], text=m["Message_Text"]) for m in msgs[:5]]
    meets = [m for m in q(c, "SELECT * FROM Meetings ORDER BY Meeting_Date DESC") if eid in [p.strip() for p in m["Participants"].split(",")]]
    meeting_rows = [dict(meeting_id=m["Meeting_ID"], topic=m["Topic"], project=pn.get(m["Project_ID"], m["Project_ID"]), date=m["Meeting_Date"],
                         decisions=m["Decisions"]) for m in meets[:5]]

    projects = {}
    for t in mine:
        if t["assignment_status"] in ("Assigned", "Completed"):
            projects.setdefault(t["project_id"], dict(project_id=t["project_id"], project=t["project"], tasks=0, open_tasks=0))
            projects[t["project_id"]]["tasks"] += 1
            projects[t["project_id"]]["open_tasks"] += 1 if t["status"] != "Completed" else 0

    # --- rule-based insights
    free = e["hours"] >= 5 and not ctx.busy.get(eid)
    ins = []
    d = compare["workload"]["department"]
    ins.append(f'Workload is {e["workload"]}%, {"above" if e["workload"] > d else "below" if e["workload"] < d else "equal to"} the {e["dept"]} average of {d}%.')
    ins.append(f'Performance score is {e["perf"]}/100, ranked {perf_rank} of {len(everyone)} employees ({"above" if e["perf"] > compare["performance"]["department"] else "at or below"} the department average of {compare["performance"]["department"]}).')
    if overdue:
        ins.append(f'{len(overdue)} open task(s) are past their deadline: ' + ", ".join(t["task_id"] for t in overdue[:5]) + ".")
    if due_soon:
        ins.append(f'{len(due_soon)} open task(s) are due within 14 days.')
    if at_risk:
        ins.append(f'{len(at_risk)} open task(s) have unresolved hard-constraint violations: ' + ", ".join(t["task_id"] for t in at_risk[:5]) + ".")
    if waiting:
        ins.append(f'{len(waiting)} received message(s) need action.')
    ins.append("Free to take a new task (at least 5 h/day and no open task)." if free else
               f'Not free for a new task: {ctx.busy.get(eid, 0)} open task(s)' + ("" if e["hours"] >= 5 else f' and only {e["hours"]} h/day available') + ".")
    if suggested:
        ins.append(f'Best pending suggestion: {suggested[0]["task_id"]} ({suggested[0]["task"]}), score {round(suggested[0]["assignment_score"])}.')

    return dict(
        generated_for=ctx.today,
        employee=dict(employee_id=eid, name=e["name"], department=e["dept"], role=e["role"], team=e["team"], work_mode=e["mode"], skills=e["skills"],
                      experience_years=e["exp"], availability_hours=e["hours"], workload_percent=e["workload"], performance_score=e["perf"], free=free),
        summary=dict(open_tasks=len(current), completed_tasks=len(finished), suggested_tasks=len(suggested), rejected_assignments=by_status.get("Rejected", 0),
                     overdue_tasks=len(overdue), due_within_14_days=len(due_soon), tasks_with_constraint_violations=len(at_risk),
                     average_assignment_score=avg(scored), projects=len(projects), messages_sent=len(sent), messages_received=len(got),
                     messages_awaiting_action=len(waiting), meetings_attended=len(meets), total_assignments=len(rows)),
        assignments_by_status=by_status, current_tasks=current, completed_tasks=finished[:10], suggested_tasks=suggested[:5],
        projects=sorted(projects.values(), key=lambda p: (-p["open_tasks"], p["project_id"])), comparison=compare,
        communications=dict(sent=len(sent), received=len(got), awaiting_action=len(waiting), recent=recent),
        meetings=dict(total=len(meets), recent=meeting_rows), insights=ins)


def meeting_summary(c, mid):
    m = q(c, "SELECT * FROM Meetings WHERE Meeting_ID=?", (mid,))
    if not m:
        raise ApiError(404, "not_found", f"Meeting {mid} not found")
    m = m[0]
    ctx = Ctx(c)
    people = [p.strip() for p in m["Participants"].split(",") if p.strip()]
    open_ = sorted((t for t in ctx.tasks.values() if t["Project_ID"] == m["Project_ID"] and t["Status"] != "Completed"), key=lambda t: t["Deadline"])[:3]
    items = [dict(task_id=t["Task_ID"], task=t["Task_Name"], responsible=(ctx.assignees(t["Task_ID"]) or [ctx.emps.get(people[0], {"name": "Unassigned"})])[0]["name"],
                  deadline=t["Deadline"], status=t["Status"]) for t in open_]
    text = (f'{m["Topic"]} with {len(people)} participants on {m["Project_ID"]}. {len(items)} open tasks need follow-up; '
            f'the next deadline is {items[0]["deadline"] if items else "none"}. Generated by rules from stored data, no language model.')
    return dict(meeting_id=mid, summary=text, decisions=m["Decisions"], action_items=items)


# ---------------------------------------------------------------- generic CRUD
def list_rows(c, route, qs):
    spec = TABLES[route]
    where, args = [], []
    for col in spec["cols"]:
        if col in qs:
            vals = qs[col][0].split(",")
            where.append(f'"{col}" IN ({",".join("?" * len(vals))})')
            args += vals
    if "q" in qs and qs["q"][0]:
        texts = [k for k, t in spec["cols"].items() if t == "text"]
        where.append("(" + " OR ".join(f'"{k}" LIKE ?' for k in texts) + ")")
        args += [f'%{qs["q"][0]}%'] * len(texts)
    sort = qs.get("sort", [spec["pk"]])[0]
    if sort not in spec["cols"]:
        raise ApiError(400, "bad_sort", f"Cannot sort by {sort}", {"allowed": list(spec["cols"])})
    order = "DESC" if qs.get("order", ["asc"])[0].lower() == "desc" else "ASC"
    try:
        size = min(1000, max(1, int(qs.get("size", ["50"])[0])))
        page = max(1, int(qs.get("page", ["1"])[0]))
    except ValueError:
        raise ApiError(400, "bad_paging", "page and size must be whole numbers")
    w = (" WHERE " + " AND ".join(where)) if where else ""
    total = c.execute(f'SELECT COUNT(*) FROM "{spec["table"]}"{w}', args).fetchone()[0]
    data = q(c, f'SELECT * FROM "{spec["table"]}"{w} ORDER BY "{sort}" {order} LIMIT ? OFFSET ?', args + [size, (page - 1) * size])
    if route == "constraints":
        with_status(c, data)
    return dict(data=data, total=total, page=page, size=size, pages=max(1, -(-total // size)))


def get_row(c, route, rid):
    spec = TABLES[route]
    r = q(c, f'SELECT * FROM "{spec["table"]}" WHERE "{spec["pk"]}"=?', (rid,))
    if not r:
        raise ApiError(404, "not_found", f"{route[:-1].capitalize()} {rid} not found")
    return with_status(c, r)[0] if route == "constraints" else r[0]


def write_row(c, route, rid, data):
    spec = TABLES[route]
    c.execute("BEGIN IMMEDIATE")
    try:
        if rid is None:
            if spec["pk"] in data and data[spec["pk"]]:
                new_id = str(data[spec["pk"]]).strip()
                if c.execute(f'SELECT 1 FROM "{spec["table"]}" WHERE "{spec["pk"]}"=?', (new_id,)).fetchone():
                    raise ApiError(409, "duplicate", f"{spec['pk']} {new_id} already exists")
            else:
                new_id = next_id(c, spec)
            body = {k: v for k, v in data.items() if k != spec["pk"]}
            row = {spec["pk"]: new_id, **validate(c, route, body, False)}
            c.execute(f'INSERT INTO "{spec["table"]}" ({",".join(chr(34) + k + chr(34) for k in row)}) VALUES ({",".join("?" * len(row))})', list(row.values()))
            rid = new_id
        else:
            cur = get_row(c, route, rid)
            cur.pop("Status", None) if route == "constraints" else None
            row = validate(c, route, data, True, cur)
            if not row:
                raise ApiError(400, "empty_update", "No fields to update")
            c.execute(f'UPDATE "{spec["table"]}" SET {",".join(chr(34) + k + chr(34) + "=?" for k in row)} WHERE "{spec["pk"]}"=?', list(row.values()) + [rid])
        c.execute("COMMIT")
    except sqlite3.IntegrityError as e:
        if c.in_transaction:
            c.execute("ROLLBACK")
        raise ApiError(409, "integrity_error", str(e))
    except BaseException:
        if c.in_transaction:
            c.execute("ROLLBACK")
        raise
    return get_row(c, route, rid)


def delete_row(c, route, rid):
    spec = TABLES[route]
    get_row(c, route, rid)
    try:
        c.execute(f'DELETE FROM "{spec["table"]}" WHERE "{spec["pk"]}"=?', (rid,))
    except sqlite3.IntegrityError:
        raise ApiError(409, "in_use", f"{rid} is referenced by other records and cannot be deleted. Remove or reassign those first.")
    return {"deleted": rid}


def bootstrap(c):
    out = {TABLES[r]["table"]: q(c, f'SELECT * FROM "{TABLES[r]["table"]}"') for r in TABLES}
    out["today"] = today()
    return out


# ---------------------------------------------------------------- routing
def route_request(method, path, qs, body):
    if body is not None and not isinstance(body, dict):
        raise ApiError(400, "bad_request", "Request body must be a JSON object")
    p = path.rstrip("/")
    if p == "/api/reports/config" and method == "GET":
        return 200, summarizer.public_config()
    if p == "/api/reports/summarize":
        if method != "POST":
            raise ApiError(405, "method_not_allowed", "Use POST")
        return 200, summarizer.summarize_upload(body["__raw"], body["__name"], qs.get("mode", ["auto"])[0], body["__ip"])
    c = connect()
    try:
        parts = [p for p in path.strip("/").split("/") if p]
        if not parts or parts[0] != "api":
            raise ApiError(404, "not_found", "Unknown API path")
        parts = parts[1:]
        if parts == ["health"]:
            return 200, dict(status="ok", today=today(), counts={r: c.execute(f'SELECT COUNT(*) FROM "{s["table"]}"').fetchone()[0] for r, s in TABLES.items()})
        if method == "GET":
            if parts == ["bootstrap"]:
                return 200, bootstrap(c)
            if parts == ["dashboard"]:
                return 200, dashboard(c)
            if parts == ["analytics"]:
                return 200, analytics(c)
            if parts == ["recommendations"]:
                return 200, recommendations(c)
            if len(parts) == 2 and parts[0] == "constraints" and parts[1] == "validate":
                tid = qs.get("task_id", [""])[0]
                ctx = Ctx(c)
                if tid not in ctx.tasks:
                    raise ApiError(404, "not_found", f"Task {tid} not found")
                rows = with_status(c, q(c, "SELECT * FROM Constraints WHERE Task_ID=?", (tid,)), ctx)
                return 200, dict(task_id=tid, total=len(rows), violated=sum(1 for r in rows if r["Status"] == "Violated"), constraints=rows)
        if parts == ["optimize"]:
            if method != "POST":
                raise ApiError(405, "method_not_allowed", "Use POST")
            try:
                top = (body or {}).get("top")
                top = None if top is None else max(1, min(150, int(top)))
                n_opt = max(1, min(10, int((body or {}).get("options", 3))))
            except (TypeError, ValueError):
                raise ApiError(422, "validation_failed", "top must be a whole number")
            return 200, optimize(c, str((body or {}).get("task_id", "")), top, n_opt)
        if parts == ["assignments", "apply"]:
            if method != "POST":
                raise ApiError(405, "method_not_allowed", "Use POST")
            return 201, apply_assignment(c, body)
        if len(parts) == 3 and parts[0] == "employees" and parts[2] == "report":
            if method != "GET":
                raise ApiError(405, "method_not_allowed", "Use GET")
            return 200, employee_report(c, unquote(parts[1]))
        if len(parts) == 3 and parts[0] == "meetings" and parts[2] == "summary":
            if method != "POST":
                raise ApiError(405, "method_not_allowed", "Use POST")
            return 200, meeting_summary(c, parts[1])
        if parts and parts[0] in TABLES:
            rt = parts[0]
            if len(parts) == 1:
                if method == "GET":
                    return 200, list_rows(c, rt, qs)
                if method == "POST":
                    return 201, write_row(c, rt, None, body)
            elif len(parts) == 2:
                rid = unquote(parts[1])
                if method == "GET":
                    return 200, get_row(c, rt, rid)
                if method in ("PUT", "PATCH"):
                    return 200, write_row(c, rt, rid, body)
                if method == "DELETE":
                    return 200, delete_row(c, rt, rid)
            raise ApiError(405, "method_not_allowed", f"{method} not supported here")
        raise ApiError(404, "not_found", "Unknown API path")
    finally:
        c.close()


class Handler(BaseHTTPRequestHandler):
    protocol_version = "HTTP/1.1"
    server_version = "ECOS/1.0"
    timeout = 60  # drop stalled or idle connections

    def log_message(self, fmt, *a):
        if os.environ.get("ECOS_QUIET") != "1":
            sys.stderr.write("%s %s\n" % (self.address_string(), fmt % a))

    def _send(self, status, payload, ctype="application/json; charset=utf-8", raw=False):
        data = payload if raw else json.dumps(payload, ensure_ascii=False).encode()
        self.send_response(status)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(data)))
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET,POST,PUT,PATCH,DELETE,OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type, X-File-Name")
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        if self.command != "HEAD":
            self.wfile.write(data)

    def _api(self, method):
        u = urlparse(self.path)
        try:
            body = None
            if method == "POST" and u.path.rstrip("/") == "/api/reports/summarize":
                # raw file upload: size is checked from the header BEFORE any bytes are read
                try:
                    n = int(self.headers.get("Content-Length") or "")
                except ValueError:
                    raise ApiError(411, "length_required", "A Content-Length header is required for uploads")
                if n > summarizer.cfg()["max_bytes"]:
                    left = min(n, 64 * summarizer.MB)  # discard the unread file so the client can read our 413 reply
                    while left > 0:
                        chunk = self.rfile.read(min(65536, left))
                        if not chunk:
                            break
                        left -= len(chunk)
                    self.close_connection = True
                    raise ApiError(413, "too_large", f'File exceeds the {round(summarizer.cfg()["max_bytes"] / summarizer.MB, 2)} MB limit.')
                body = {"__raw": self.rfile.read(n) if n else b"", "__name": unquote(self.headers.get("X-File-Name", "")), "__ip": self.client_address[0]}
            elif method in ("POST", "PUT", "PATCH"):
                n = int(self.headers.get("Content-Length") or 0)
                if n > MAX_BODY:
                    self.close_connection = True
                    raise ApiError(413, "too_large", "Request body too large")
                raw = self.rfile.read(n) if n else b""
                if raw.strip():
                    try:
                        body = json.loads(raw)
                    except ValueError:
                        raise ApiError(400, "invalid_json", "Request body is not valid JSON")
                else:
                    body = {}
            status, payload = route_request(method, u.path, parse_qs(u.query), body)
        except ApiError as e:
            status, payload = e.status, {"error": {"code": e.code, "message": e.message, "details": e.details}}
        except summarizer.SummaryError as e:
            status, payload = e.status, {"error": {"code": e.code, "message": e.message}}
        except sqlite3.OperationalError as e:
            status, payload = 503, {"error": {"code": "database_error", "message": f"Database unavailable: {e}. Run import_excel.py first?"}}
        except Exception as e:  # never leak a traceback to the client
            sys.stderr.write(f"Unhandled error: {e!r}\n")
            status, payload = 500, {"error": {"code": "internal_error", "message": "Unexpected server error"}}
        self._send(status, payload)

    def _static(self):
        p = unquote(urlparse(self.path).path)
        full = os.path.abspath(os.path.join(FRONT, "index.html" if p in ("", "/") else p.lstrip("/")))
        if not full.startswith(FRONT + os.sep) or not os.path.isfile(full):
            return self._send(404, {"error": {"code": "not_found", "message": "File not found"}})
        with open(full, "rb") as f:
            self._send(200, f.read(), mimetypes.guess_type(full)[0] or "application/octet-stream", raw=True)

    def do_OPTIONS(self):
        self._send(204, b"", raw=True)

    def do_HEAD(self):
        self._static()

    def do_GET(self):
        self._api("GET") if self.path.startswith("/api") else self._static()

    def do_POST(self):
        self._api("POST")

    def do_PUT(self):
        self._api("PUT")

    def do_PATCH(self):
        self._api("PATCH")

    def do_DELETE(self):
        self._api("DELETE")


def main():
    if not os.path.exists(DB):
        sys.exit(f"Database not found at {DB}. Run: python import_excel.py")
    
host, port = os.environ.get("HOST", "0.0.0.0"), int(os.environ.get("PORT") or (sys.argv[1] if len(sys.argv) > 1 else 8000))

    srv = ThreadingHTTPServer((host, port), Handler)
    print(f"ECOS running at http://{'localhost' if host == '127.0.0.1' else host}:{port}  (API under /api, database {DB})")
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        print("\nStopped.")


if __name__ == "__main__":
    main()
