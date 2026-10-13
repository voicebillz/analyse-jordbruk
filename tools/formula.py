# -*- coding: utf-8 -*-
"""Omset formelspråket i spec.py til Excel-formlar, JavaScript og forklarande tekst."""
import re

TOKEN = re.compile(r'\s*(?:(\d+\.\d+|\d+)|("[^"]*")|([A-Za-z_][A-Za-z0-9_]*(?:\.[A-Za-z_][A-Za-z0-9_]*)?)|(<=|>=|<>|[-+*/(),<>=]))')
FUNCS = {"DIV", "IF", "MAX", "MIN", "ABS", "ROUND", "ACTIVE", "SUM"}


def tokenize(expr):
    pos, out = 0, []
    expr = expr.strip()
    while pos < len(expr):
        m = TOKEN.match(expr, pos)
        if not m or m.end() == pos:
            raise ValueError("Kan ikkje tolke formel: %r ved %d" % (expr, pos))
        num, s, ident, op = m.groups()
        if num is not None:
            out.append(("num", num))
        elif s is not None:
            out.append(("str", s))
        elif ident is not None:
            out.append(("fn" if ident in FUNCS else "id", ident))
        else:
            out.append(("op", op))
        pos = m.end()
    return out


def _split_call(tokens, i):
    """tokens[i] er funksjonsnamn, tokens[i+1] er '('. Returnerer (args, index etter ')')."""
    assert tokens[i + 1] == ("op", "(")
    depth, j, args, cur = 0, i + 1, [], []
    while True:
        t = tokens[j]
        if t == ("op", "("):
            depth += 1
            if depth > 1:
                cur.append(t)
        elif t == ("op", ")"):
            depth -= 1
            if depth == 0:
                args.append(cur)
                return args, j + 1
            cur.append(t)
        elif t == ("op", ",") and depth == 1:
            args.append(cur)
            cur = []
        else:
            cur.append(t)
        j += 1


def _simple(tokens):
    return len(tokens) == 1


def render(tokens, ref, active, mode):
    """mode: 'xl' | 'js' | 'txt'"""
    out, i = [], 0
    while i < len(tokens):
        kind, v = tokens[i]
        if kind == "fn":
            args, j = _split_call(tokens, i)
            if v == "ACTIVE":
                key = args[0][0][1]
                out.append(active(key))
            else:
                ra = [render(a, ref, active, mode) for a in args]
                if v == "DIV":
                    if mode == "xl":
                        out.append("IF((%s)=0,0,(%s)/(%s))" % (ra[1], ra[0], ra[1]))
                    elif mode == "js":
                        out.append("DIV(%s,%s)" % (ra[0], ra[1]))
                    else:
                        a = ra[0] if _simple(args[0]) else "(%s)" % ra[0]
                        b = ra[1] if _simple(args[1]) else "(%s)" % ra[1]
                        out.append("%s ÷ %s" % (a, b))
                else:
                    if mode == "txt":
                        name = {"IF": "HVIS", "MAX": "MAKS", "MIN": "MIN", "ABS": "ABS", "ROUND": "AVRUND", "SUM": "SUM"}[v]
                        out.append("%s(%s)" % (name, "; ".join(ra)))
                    else:
                        out.append("%s(%s)" % (v, ",".join(ra)))
            i = j
            continue
        if kind == "id":
            out.append(ref(v))
        elif kind == "num":
            out.append(v.replace(".", ",") if mode == "txt" else v)
        elif kind == "str":
            out.append(v)
        else:
            if mode == "txt":
                out.append({"*": " × ", "-": " − ", "+": " + ", "/": " / ", ",": "; "}.get(v, v))
            elif mode == "js" and v == "=":
                out.append("==")
            else:
                out.append(v)
        i += 1
    return "".join(out)


def ids_in(expr):
    return [v for k, v in tokenize(expr) if k == "id"]
