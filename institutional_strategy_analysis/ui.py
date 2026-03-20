# -*- coding: utf-8 -*-
"""
institutional_strategy_analysis/ui.py
───────────────────────────────────────
Self-contained Streamlit UI for "ניתוח אסטרטגיות מוסדיים".
Renders as an st.expander at the bottom of the main app.

Entry point (one line in streamlit_app.py):
    from institutional_strategy_analysis.ui import render_institutional_analysis
    render_institutional_analysis()

All session-state keys are prefixed "isa_" to avoid any collision.
"""
from __future__ import annotations

from datetime import date, timedelta

import pandas as pd
import streamlit as st

# ── Sheet URL ─────────────────────────────────────────────────────────────────
# ▼▼▼  Set your Google Sheets URL here  ▼▼▼
ISA_SHEET_URL = (
    "https://docs.google.com/spreadsheets/d/"
    "1e9zjj1OWMYqUYoK6YFYvYwOnN7qbydYDyArHbn8l9pE/edit"
)
# ▲▲▲─────────────────────────────────────────────────────────────────────────

# ── Lazy imports (never execute at import time) ───────────────────────────────

def _load_data():
    from institutional_strategy_analysis.loader     import load_raw_blocks
    from institutional_strategy_analysis.series_builder import get_time_bounds
    import streamlit as st

    @st.cache_data(ttl=3600, show_spinner=False)
    def _cached(url: str):
        return load_raw_blocks(url)

    return _cached(ISA_SHEET_URL)


def _build_series(df_y, df_m, rng, custom_start, filters):
    from institutional_strategy_analysis.series_builder import build_display_series
    return build_display_series(df_y, df_m, rng, custom_start, filters)


def _options(df_y, df_m):
    from institutional_strategy_analysis.series_builder import get_available_options
    return get_available_options(df_y, df_m)


# ── Helpers ───────────────────────────────────────────────────────────────────

def _safe_plotly(fig, key=None):
    try:
        st.plotly_chart(fig, use_container_width=True, key=key)
    except TypeError:
        st.plotly_chart(fig)


def _csv(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")


def _clamp(val: date, lo: date, hi: date) -> date:
    return max(lo, min(hi, val))


# ── Debug panel ───────────────────────────────────────────────────────────────

def _render_debug(df_yearly, df_monthly, debug_info, errors):
    with st.expander("🛠️ מידע אבחון (debug)", expanded=False):
        col1, col2 = st.columns(2)
        with col1:
            st.metric("גליונות שנטענו", len(debug_info))
            st.metric("שורות שנתי", len(df_yearly))
            st.metric("שורות חודשי", len(df_monthly))
        with col2:
            if not df_yearly.empty:
                yr = df_yearly["date"]
                st.metric("טווח שנתי", f"{yr.min().year} – {yr.max().year}")
            if not df_monthly.empty:
                mr = df_monthly["date"]
                st.metric("טווח חודשי",
                          f"{mr.min().strftime('%Y-%m')} – {mr.max().strftime('%Y-%m')}")

        if debug_info:
            rows = []
            for d in debug_info:
                rows.append({
                    "גליון": d.get("sheet", "?"),
                    "header row": d.get("header_row", "?"),
                    "freq col": d.get("freq_col", "—"),
                    "שורות שנתיות": d.get("yearly_rows", 0),
                    "שורות חודשיות": d.get("monthly_rows", 0),
                    "טווח שנתי": d.get("yearly_range", "—"),
                    "טווח חודשי": d.get("monthly_range", "—"),
                    "שגיאה": d.get("error", ""),
                })
            st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

        if errors:
            for e in errors:
                st.warning(e)



# ── AI Analysis renderer ──────────────────────────────────────────────────────

_SECTION_ICONS = {
    # Market analysis
    "מיצוב יחסי לפי רכיב":                           "🎯",
    "דינמיות ואפיון סגנון ניהול":                    "⚡",
    'אסטרטגיית גידור מט"ח':                          "🛡️",
    "תנועות פוזיציה אחרונות (3–12 חודשים)":          "🔄",
    "ניתוח סיכון קבוצתי":                            "⚠️",
    "יתרונות וחסרונות לפי גוף":                      "✅",
    "תובנה אסטרטגית וסיכום":                         "💡",
    # Focused analysis
    "מיצוב יחסי (Relative Positioning)":             "🎯",
    "ניתוח היסטורי עצמי (Historical Self-Analysis)": "📜",
    "סגנון ניהול ועקביות (Management Style & Consistency)": "⚡",
    "אסטרטגיית גידור ומטבע (Hedging & Currency Strategy)": "🛡️",
    "מומנטום ותנועות אחרונות (Recent Momentum)":     "🔄",
    "פרופיל סיכון (Risk Assessment)":                "⚠️",
    "גזר דין — בחירת מנהל (Manager Selection Verdict)": "👑",
    # Comparison
    "סיכום מנהלי (Executive Summary)":               "📌",
    "השוואה יחסית לפי רכיב":                        "⚖️",
    "הבדלי סגנון ניהול (Management Style Delta)":    "⚡",
    "אסטרטגיית גידור (Hedging Comparison)":         "🛡️",
    "המלצה לפי פרופיל משקיע":                       "👤",
    # Legacy fallbacks
    "ניתוח לפי גוף ומסלול":    "🏢",
    "ניתוח סיכון":              "⚠️",
    "תובנה אסטרטגית":          "💡",
    "סיכום מנהלי":             "📌",
}

_EXPANDED_SECTIONS = {
    "מיצוב יחסי",
    "סיכום מנהלי",
    "גזר דין",
    "תובנה אסטרטגית",
}


def _render_api_key_input():
    """Returns True if an OpenAI API key is available (from Secrets or session)."""
    # Check Secrets first — key pre-configured by admin
    try:
        if hasattr(st, "secrets") and "OPENAI_API_KEY" in st.secrets:
            return True
    except Exception:
        pass

    existing = st.session_state.get("isa_api_key", "").strip()
    if existing:
        col_ok, col_clr = st.columns([5, 1])
        with col_ok:
            st.success("✅ מפתח OpenAI API מוגדר", icon="🔑")
        with col_clr:
            if st.button("נקה", key="isa_api_key_clear", help="הסר מפתח"):
                st.session_state["isa_api_key"] = ""
                st.rerun()
        return True

    with st.expander("🔑 הגדרת מפתח OpenAI API לניתוח AI", expanded=True):
        st.markdown(
            "כדי להפעיל ניתוח AI יש להזין מפתח [OpenAI API](https://platform.openai.com/api-keys). "
            "המפתח נשמר בזיכרון הדפדפן בלבד ואינו מועלה לשום שרת."
        )
        key_input = st.text_input(
            "OPENAI_API_KEY",
            type="password",
            placeholder="sk-...",
            key="isa_api_key_input_field",
            label_visibility="collapsed",
        )
        if st.button("אשר מפתח", key="isa_api_key_confirm", type="primary"):
            if key_input.strip().startswith("sk-"):
                st.session_state["isa_api_key"] = key_input.strip()
                st.rerun()
            else:
                st.error("מפתח לא תקין — חייב להתחיל ב-sk-...")
    return False


def _render_analysis_result(result, cache_key: str, dl_key: str, refresh_key: str):
    """Display an AnalysisResult with styled section expanders."""
    if result.error:
        st.error(f"⚠️ {result.error}")
        if st.button("נסה שוב", key=f"{refresh_key}_retry_{cache_key}"):
            st.session_state.pop(cache_key, None)
            st.session_state.pop(f"{cache_key}_sig", None)
            st.rerun()
        return

    if not result.sections:
        st.markdown(result.raw_text)
    else:
        for title, body in result.sections.items():
            if title == "כללי" and not body.strip():
                continue
            icon = _SECTION_ICONS.get(title, "📋")
            exp  = any(s in title for s in _EXPANDED_SECTIONS)
            with st.expander(f"{icon} {title}", expanded=exp):
                st.markdown(body)

    col_dl, col_rf, _ = st.columns([1, 1, 4])
    with col_dl:
        st.download_button(
            "⬇️ שמור ניתוח",
            data=result.raw_text.encode("utf-8"),
            file_name=f"{dl_key}.txt",
            mime="text/plain",
            key=f"isa_dl_{dl_key}_{cache_key}",
            use_container_width=True,
        )
    with col_rf:
        if st.button("🔄 רענן", key=f"{refresh_key}_{cache_key}",
                     help="הרץ מחדש את הניתוח", use_container_width=True):
            st.session_state.pop(cache_key, None)
            st.session_state.pop(f"{cache_key}_sig", None)
            st.rerun()


def _scorecard_badge(diff: float) -> str:
    """Return an HTML badge for relative positioning."""
    if diff > 3:   return "<span style='background:#16a34a;color:#fff;padding:2px 8px;border-radius:99px;font-size:11px;font-weight:700'>▲ גבוה משמעותית</span>"
    if diff > 1:   return "<span style='background:#4ade80;color:#14532d;padding:2px 8px;border-radius:99px;font-size:11px;font-weight:700'>▲ מעל ממוצע</span>"
    if diff < -3:  return "<span style='background:#dc2626;color:#fff;padding:2px 8px;border-radius:99px;font-size:11px;font-weight:700'>▼ נמוך משמעותית</span>"
    if diff < -1:  return "<span style='background:#f87171;color:#7f1d1d;padding:2px 8px;border-radius:99px;font-size:11px;font-weight:700'>▼ מתחת לממוצע</span>"
    return "<span style='background:#e2e8f0;color:#475569;padding:2px 8px;border-radius:99px;font-size:11px;font-weight:700'>◼ ממוצע</span>"


def _direction_badge(direction: str) -> str:
    if direction == "עולה":   return "🟢 עולה"
    if direction == "יורדת":  return "🔴 יורדת"
    return "⚪ יציבה"


def _render_quick_scorecard(full_df: pd.DataFrame, manager: str, track: str):
    """Render a quick stat-based scorecard card before running AI."""
    import numpy as np
    try:
        from institutional_strategy_analysis.ai_analyst import compute_manager_scorecard
        rows = compute_manager_scorecard(full_df, manager, track)
    except Exception:
        return

    if not rows:
        return

    st.markdown("""
<div style='background:#f0f4ff;border:1px solid #c7d7fe;border-radius:10px;
     padding:14px 18px;margin:10px 0 4px 0;direction:rtl'>
  <div style='font-size:13px;font-weight:700;color:#1e3a8a;margin-bottom:10px'>
    📊 סקירה מהירה — מיצוב יחסי לפי רכיב
  </div>""", unsafe_allow_html=True)

    cols = st.columns(len(rows))
    for col, row in zip(cols, rows):
        diff = row["diff_mean"]
        c3   = row.get("change_3m", float("nan"))
        c12  = row.get("change_12m", float("nan"))
        import math
        c3s  = f"{c3:+.1f}pp" if not math.isnan(c3)  else "—"
        c12s = f"{c12:+.1f}pp" if not math.isnan(c12) else "—"
        with col:
            st.markdown(f"""
<div style='background:#fff;border:1px solid #e2e8f0;border-radius:8px;
     padding:10px;text-align:center;direction:rtl'>
  <div style='font-size:11px;color:#64748b;margin-bottom:4px'>{row['alloc']}</div>
  <div style='font-size:22px;font-weight:900;color:#1e3a8a'>{row['current']}%</div>
  <div style='font-size:11px;color:#64748b'>ממוצע קבוצה: {row['peer_mean']}%</div>
  <div style='font-size:11px;margin:4px 0'>{_scorecard_badge(diff)}</div>
  <div style='font-size:10px;color:#94a3b8;margin-top:4px'>
    3ח: {c3s} | 12ח: {c12s}<br/>
    דירוג: {row['rank']}/{row['n_total']} | {_direction_badge(row['direction'])}
  </div>
</div>""", unsafe_allow_html=True)

    st.markdown("</div>", unsafe_allow_html=True)


def _render_ai_section(
    display_df: pd.DataFrame,
    full_df: pd.DataFrame,
    context: dict,
    sel_mgr: list,
    sel_tracks: list,
):
    """
    Professional 3-mode AI analysis panel:
      MODE 1 — ניתוח שוק — comparative analysis across all selected managers
      MODE 2 — ניתוח מיקוד — deep relative analysis of ONE manager vs peer group
      MODE 3 — השוואה ישירה — head-to-head A vs B
    """
    st.markdown("---")
    st.markdown("""
<div style='background:linear-gradient(135deg,#0f2657 0%,#1d4ed8 60%,#2563eb 100%);
     border-radius:14px;padding:18px 24px;margin-bottom:18px;direction:rtl'>
  <div style='color:#fff;font-size:20px;font-weight:900;letter-spacing:-0.5px'>
    🤖 מנוע ניתוח AI — אסטרטגיות מוסדיים
  </div>
  <div style='color:#93c5fd;font-size:12px;margin-top:5px;line-height:1.6'>
    מיצוב יחסי · היסטוריה עצמית · גידור מט"ח · סגנון ניהול · בחירת מנהל
  </div>
  <div style='display:flex;gap:8px;margin-top:10px;flex-wrap:wrap'>
    <span style='background:rgba(255,255,255,.15);color:#e0eaff;padding:3px 10px;
      border-radius:99px;font-size:11px'>CIO Level Analysis</span>
    <span style='background:rgba(255,255,255,.15);color:#e0eaff;padding:3px 10px;
      border-radius:99px;font-size:11px'>Manager Selection</span>
    <span style='background:rgba(255,255,255,.15);color:#e0eaff;padding:3px 10px;
      border-radius:99px;font-size:11px'>Relative Positioning</span>
  </div>
</div>""", unsafe_allow_html=True)

    # ── API key ───────────────────────────────────────────────────────────
    has_key = _render_api_key_input()
    if not has_key:
        return

    # ── Mode selector ─────────────────────────────────────────────────────
    mode_labels = {
        "market":    "🌐 ניתוח שוק — השוואה בין גופים",
        "focused":   "🎯 ניתוח מיקוד — גוף אחד לעומת עמיתים",
        "headtohead": "⚖️ השוואה ישירה — A מול B",
    }
    mode_keys   = list(mode_labels.keys())
    mode_idx    = st.session_state.get("isa_ai_mode_idx", 0)

    mc1, mc2, mc3 = st.columns(3)
    for col, (k, label), idx in zip(
        [mc1, mc2, mc3], mode_labels.items(), range(3)
    ):
        selected = (mode_idx == idx)
        bg  = "linear-gradient(135deg,#1d4ed8,#2563eb)" if selected else "#f8faff"
        txt = "#ffffff" if selected else "#1e3a8a"
        brd = "#1d4ed8" if selected else "#c7d7fe"
        with col:
            st.markdown(f"""
<div style='background:{bg};border:2px solid {brd};border-radius:10px;
     padding:10px 12px;text-align:center;cursor:pointer;
     direction:rtl;color:{txt};font-size:13px;font-weight:600;margin-bottom:6px'>
  {label}
</div>""", unsafe_allow_html=True)
            if st.button(
                "✓ בחר" if selected else "בחר",
                key=f"isa_mode_{k}",
                type="primary" if selected else "secondary",
                use_container_width=True,
            ):
                st.session_state["isa_ai_mode_idx"] = idx
                st.rerun()

    st.markdown("<hr style='border:none;border-top:1px solid #e2e8f0;margin:10px 0'>",
                unsafe_allow_html=True)

    # Get combos available in display_df
    combos = sorted(
        display_df[["manager", "track"]].drop_duplicates()
        .apply(lambda r: f"{r['manager']} | {r['track']}", axis=1)
        .tolist()
    )
    all_mgrs = sorted(full_df["manager"].unique().tolist()) if not full_df.empty else sel_mgr

    # ══════════════════════════════════════════════════════════════════════
    # MODE 1 — Market Analysis
    # ══════════════════════════════════════════════════════════════════════
    if mode_idx == 0:
        st.markdown("""
<div style='background:#f0f9ff;border-right:4px solid #1d4ed8;padding:10px 16px;
     border-radius:0 8px 8px 0;margin-bottom:14px;direction:rtl'>
  <b style='color:#1e3a8a'>ניתוח שוק</b>
  <span style='color:#475569;font-size:12px'> — ניתוח יחסי בין כל הגופים שנבחרו: מיצוב, דינמיות, גידור מט"ח, סיכון</span>
</div>""", unsafe_allow_html=True)

        cache_key  = "isa_market_result"
        filter_sig = str(sorted((k, str(v)) for k, v in context.items()))
        if st.session_state.get("isa_market_sig") != filter_sig:
            st.session_state.pop(cache_key, None)
            st.session_state["isa_market_sig"] = filter_sig

        if cache_key not in st.session_state:
            st.caption(f"גופים בניתוח: {', '.join(sel_mgr)} | מסלולים: {', '.join(sel_tracks)}")
            if st.button("🌐 הפעל ניתוח שוק", key="isa_market_btn", type="primary",
                         use_container_width=False):
                with st.spinner("Claude מנתח השוואה בין גופים... (עד 90 שניות)"):
                    try:
                        from institutional_strategy_analysis.ai_analyst import run_ai_analysis
                        result = run_ai_analysis(display_df, context)
                        st.session_state[cache_key] = result
                    except Exception as e:
                        st.error(f"שגיאה: {e}")
                        return
                st.rerun()
        else:
            st.caption(f"גופים: {', '.join(sel_mgr)} | מסלולים: {', '.join(sel_tracks)}")
            _render_analysis_result(
                st.session_state[cache_key],
                cache_key, "market_analysis", "isa_market_refresh",
            )

    # ══════════════════════════════════════════════════════════════════════
    # MODE 2 — Focused Single-Manager Analysis
    # ══════════════════════════════════════════════════════════════════════
    elif mode_idx == 1:
        st.markdown("""
<div style='background:#f0fdf4;border-right:4px solid #16a34a;padding:10px 16px;
     border-radius:0 8px 8px 0;margin-bottom:14px;direction:rtl'>
  <b style='color:#14532d'>ניתוח מיקוד</b>
  <span style='color:#475569;font-size:12px'> — גוף אחד לעומת כל עמיתיו: מיצוב יחסי, היסטוריה עצמית, סגנון ניהול, גזר דין</span>
</div>""", unsafe_allow_html=True)

        # Select one manager + track
        f1, f2 = st.columns(2)
        with f1:
            focus_mgr = st.selectbox(
                "גוף מנהל לניתוח",
                options=all_mgrs,
                index=0,
                key="isa_focus_mgr",
                help="הגוף שיינתח לעומק מול כלל עמיתיו",
            )
        avail_tracks_focus = sorted(
            full_df[full_df["manager"] == focus_mgr]["track"].unique().tolist()
        ) if not full_df.empty else sel_tracks
        with f2:
            focus_trk = st.selectbox(
                "מסלול",
                options=avail_tracks_focus,
                index=0,
                key="isa_focus_trk",
                help="המסלול שיינתח",
            )

        # Optional: limit peer group
        peer_options = [m for m in all_mgrs if m != focus_mgr]
        use_custom_peers = st.toggle(
            "הגבל קבוצת ייחוס לגופים ספציפיים",
            value=False,
            key="isa_custom_peers_toggle",
            help="כברירת מחדל — כל הגופים האחרים. הפעל כדי לבחור עמיתים ספציפיים.",
        )
        custom_peers = None
        if use_custom_peers and peer_options:
            custom_peers = st.multiselect(
                "גופי ייחוס (peer group)",
                options=peer_options,
                default=peer_options[:min(4, len(peer_options))],
                key="isa_focus_peers",
                help="הגופים שישמשו להשוואה יחסית",
            )
            if not custom_peers:
                st.warning("יש לבחור לפחות גוף אחד לקבוצת הייחוס.")
                return

        # Quick scorecard (no API)
        _render_quick_scorecard(full_df if not full_df.empty else display_df, focus_mgr, focus_trk)

        # Build cache key
        peer_str  = "|".join(sorted(custom_peers)) if custom_peers else "all"
        cache_key = f"isa_focus_{focus_mgr}_{focus_trk}_{peer_str}"
        cache_key = cache_key.replace(" ", "_")[:80]

        if cache_key not in st.session_state:
            peers_display = ", ".join(custom_peers) if custom_peers else f"כל הגופים ({len(peer_options)})"
            st.caption(f"קבוצת ייחוס: {peers_display}")
            if st.button(
                f"🎯 הפעל ניתוח מיקוד — {focus_mgr}",
                key="isa_focus_btn", type="primary",
            ):
                with st.spinner(f"Claude מנתח {focus_mgr} לעומק... (עד 90 שניות)"):
                    try:
                        from institutional_strategy_analysis.ai_analyst import run_focused_analysis
                        fd = full_df if not full_df.empty else display_df
                        focused_ctx = {**context,
                                       "selected_manager": focus_mgr,
                                       "selected_track":   focus_trk}
                        result = run_focused_analysis(
                            fd, focus_mgr, focus_trk, custom_peers, focused_ctx
                        )
                        st.session_state[cache_key] = result
                    except Exception as e:
                        st.error(f"שגיאה: {e}")
                        return
                st.rerun()
        else:
            st.markdown(f"""
<div style='background:#f0fdf4;border:1px solid #86efac;border-radius:8px;
     padding:8px 14px;margin-bottom:8px;direction:rtl;font-size:13px;color:#14532d'>
  📋 ניתוח מיקוד: <b>{focus_mgr}</b> | מסלול: <b>{focus_trk}</b>
</div>""", unsafe_allow_html=True)
            _render_analysis_result(
                st.session_state[cache_key],
                cache_key, f"focused_{focus_mgr[:8]}", "isa_focus_refresh",
            )

    # ══════════════════════════════════════════════════════════════════════
    # MODE 3 — Head-to-Head Comparison
    # ══════════════════════════════════════════════════════════════════════
    else:
        st.markdown("""
<div style='background:#fff7ed;border-right:4px solid #ea580c;padding:10px 16px;
     border-radius:0 8px 8px 0;margin-bottom:14px;direction:rtl'>
  <b style='color:#7c2d12'>השוואה ישירה</b>
  <span style='color:#475569;font-size:12px'> — השוואת A מול B: יתרונות, חסרונות, המלצה לפי פרופיל משקיע</span>
</div>""", unsafe_allow_html=True)

        if len(combos) < 2:
            st.info("יש לוודא שנבחרו לפחות 2 גופים בסינון הנתונים.")
            return

        cc1, cc2 = st.columns(2)
        with cc1:
            st.markdown("""<div style='background:#dbeafe;border-radius:8px 8px 0 0;
                padding:6px 12px;font-size:12px;font-weight:700;color:#1e40af;
                text-align:center;direction:rtl'>גוף A</div>""", unsafe_allow_html=True)
            combo_a = st.selectbox(
                "גוף A", options=combos, index=0, key="isa_cmp_a",
                label_visibility="collapsed",
            )
        with cc2:
            st.markdown("""<div style='background:#fef3c7;border-radius:8px 8px 0 0;
                padding:6px 12px;font-size:12px;font-weight:700;color:#92400e;
                text-align:center;direction:rtl'>גוף B</div>""", unsafe_allow_html=True)
            default_b = 1 if len(combos) > 1 else 0
            combo_b = st.selectbox(
                "גוף B", options=combos, index=default_b, key="isa_cmp_b",
                label_visibility="collapsed",
            )

        # Visual diff preview
        if combo_a != combo_b:
            mgr_a_p, trk_a_p = combo_a.split(" | ", 1)
            mgr_b_p, trk_b_p = combo_b.split(" | ", 1)
            import numpy as np
            try:
                from institutional_strategy_analysis.ai_analyst import _compute_manager_profile
                pa = _compute_manager_profile(display_df, mgr_a_p, trk_a_p)
                pb = _compute_manager_profile(display_df, mgr_b_p, trk_b_p)
                if pa and pb:
                    shared = sorted(set(pa["per_alloc"]) & set(pb["per_alloc"]))
                    if shared:
                        preview_rows = []
                        for alloc in shared:
                            sa = pa["per_alloc"][alloc]
                            sb = pb["per_alloc"][alloc]
                            diff = sa["current"] - sb["current"]
                            preview_rows.append({
                                "רכיב": alloc,
                                f"{mgr_a_p}": f"{sa['current']}%",
                                f"{mgr_b_p}": f"{sb['current']}%",
                                "הפרש": f"{diff:+.1f}pp",
                                "יותר תזזיתי": mgr_a_p if sa['dynamism'] > sb['dynamism'] else mgr_b_p,
                            })
                        if preview_rows:
                            import pandas as pd
                            st.dataframe(
                                pd.DataFrame(preview_rows),
                                use_container_width=True,
                                hide_index=True,
                            )
            except Exception:
                pass

        cmp_cache = f"isa_cmp_{combo_a}_{combo_b}".replace(" ", "_").replace("|", "_")[:80]
        cmp_sig   = f"{combo_a}|{combo_b}"

        if st.session_state.get("isa_cmp_sig") != cmp_sig:
            for k in list(st.session_state.keys()):
                if k.startswith("isa_cmp_") and k not in ("isa_cmp_a", "isa_cmp_b", "isa_cmp_sig"):
                    st.session_state.pop(k, None)
            st.session_state["isa_cmp_sig"] = cmp_sig

        if cmp_cache not in st.session_state:
            if combo_a == combo_b:
                st.warning("יש לבחור שני גופים/מסלולים שונים.")
            else:
                if st.button("⚖️ הפעל השוואה", key="isa_cmp_btn", type="primary"):
                    mgr_a, trk_a = combo_a.split(" | ", 1)
                    mgr_b, trk_b = combo_b.split(" | ", 1)
                    with st.spinner(f"משווה {mgr_a} מול {mgr_b}... (עד 90 שניות)"):
                        try:
                            from institutional_strategy_analysis.ai_analyst import run_comparison_analysis
                            result = run_comparison_analysis(
                                display_df, mgr_a, trk_a, mgr_b, trk_b, context
                            )
                            st.session_state[cmp_cache] = result
                        except Exception as e:
                            st.error(f"שגיאה: {e}")
                            return
                    st.rerun()
        else:
            mgr_a_l, _ = combo_a.split(" | ", 1)
            mgr_b_l, _ = combo_b.split(" | ", 1)
            st.markdown(f"""
<div style='display:flex;gap:10px;align-items:center;justify-content:center;
     margin:6px 0 12px 0;direction:rtl'>
  <span style='background:#dbeafe;color:#1e40af;font-weight:700;padding:4px 14px;
    border-radius:8px;font-size:13px'>{mgr_a_l}</span>
  <span style='color:#64748b;font-size:14px'>⚔️</span>
  <span style='background:#fef3c7;color:#92400e;font-weight:700;padding:4px 14px;
    border-radius:8px;font-size:13px'>{mgr_b_l}</span>
</div>""", unsafe_allow_html=True)
            _render_analysis_result(
                st.session_state[cmp_cache],
                cmp_cache, f"cmp_{mgr_a_l[:6]}_{mgr_b_l[:6]}", "isa_cmp_refresh",
            )


# ── Main entry point ──────────────────────────────────────────────────────────

def render_institutional_analysis():
    """Render the full "ניתוח אסטרטגיות מוסדיים" section."""

    with st.expander("📐 ניתוח אסטרטגיות מוסדיים", expanded=False):

        # ── Load data ─────────────────────────────────────────────────────
        with st.spinner("טוען נתונים..."):
            try:
                df_yearly, df_monthly, debug_info, errors = _load_data()
            except Exception as e:
                st.error(f"שגיאת טעינה: {e}")
                return

        if df_yearly.empty and df_monthly.empty:
            st.error("לא נטענו נתונים. בדוק את קישור הגיליון ואת הרשאות הגישה.")
            for e in errors:
                st.warning(e)
            return

        _render_debug(df_yearly, df_monthly, debug_info, errors)

        # ── Available options ─────────────────────────────────────────────
        opts = _options(df_yearly, df_monthly)

        # ── Filters ───────────────────────────────────────────────────────
        st.markdown("#### 🎛️ סינון")
        fc1, fc2, fc3 = st.columns(3)

        preferred_mgrs = [m for m in ["הראל", "מגדל"] if m in opts["managers"]]
        default_mgrs = preferred_mgrs or opts["managers"][:min(2, len(opts["managers"]))]

        with fc1:
            sel_mgr = st.multiselect(
                "מנהל השקעות",
                options=opts["managers"],
                default=default_mgrs,
                help="בחר גוף מוסדי אחד או יותר. הנתונים מציגים את אסטרטגיית האלוקציה שלהם לאורך זמן.",
                key="isa_managers",
            )
        with fc2:
            avail_tracks = sorted({
                t for df in (df_yearly, df_monthly) if not df.empty
                for t in df[df["manager"].isin(sel_mgr)]["track"].unique()
            }) if sel_mgr else opts["tracks"]
            default_tracks = [t for t in ["כללי"] if t in avail_tracks] or (avail_tracks[:1] if avail_tracks else [])
            sel_tracks = st.multiselect(
                "מסלול",
                options=avail_tracks,
                default=default_tracks,
                help="בחר מסלול השקעה — כגון כללי, מנייתי. מסלול כללי מאזן בין כמה נכסים.",
                key="isa_tracks",
            )
        with fc3:
            avail_allocs = sorted({
                a for df in (df_yearly, df_monthly) if not df.empty
                for a in df[
                    df["manager"].isin(sel_mgr) & df["track"].isin(sel_tracks)
                ]["allocation_name"].unique()
            }) if sel_mgr and sel_tracks else opts["allocation_names"]
            default_allocs = [a for a in avail_allocs if a == 'חו"ל']
            if not default_allocs:
                default_allocs = [a for a in avail_allocs if "חו" in a or "חול" in a][:1]
            if not default_allocs:
                default_allocs = avail_allocs[:1] if avail_allocs else []
            sel_allocs = st.multiselect(
                "רכיב אלוקציה",
                options=avail_allocs,
                default=default_allocs,
                help='בחר רכיבי חשיפה — למשל מניות, חו"ל, מט"ח, לא-סחיר.',
                key="isa_allocs",
            )

        # Time range
        rng_c, cust_c = st.columns([3, 2])
        with rng_c:
            sel_range = st.radio(
                "טווח זמן",
                options=["הכל", "YTD", "1Y", "3Y", "5Y", "מותאם אישית"],
                index=0, horizontal=True,
                label_visibility="collapsed",
                key="isa_range",
            )
            st.caption(
                "⏱️ **טווח זמן** — YTD ו-1Y משתמשים בנתונים חודשיים בלבד. "
                "3Y/5Y/הכל משלבים חודשי + שנתי."
            )
        with cust_c:
            custom_start = None
            if sel_range == "מותאם אישית":
                from institutional_strategy_analysis.series_builder import get_time_bounds
                min_d, max_d = get_time_bounds(df_yearly, df_monthly)
                custom_start = st.date_input(
                    "מתאריך", value=min_d.date(),
                    min_value=min_d.date(), max_value=max_d.date(),
                    key="isa_custom_start",
                )

        if not sel_mgr or not sel_tracks or not sel_allocs:
            st.info("יש לבחור לפחות מנהל, מסלול ורכיב אחד.")
            return

        # ── Build display series ──────────────────────────────────────────
        filters = {"managers": sel_mgr, "tracks": sel_tracks,
                   "allocation_names": sel_allocs}

        display_df = _build_series(df_yearly, df_monthly, sel_range, custom_start, filters)

        if display_df.empty:
            if sel_range in ("YTD", "1Y") and df_monthly.empty:
                st.warning(
                    "⚠️ לא נמצאו נתונים חודשיים. "
                    "YTD ו-1Y דורשים נתונים חודשיים. "
                    "נסה 'הכל' או '3Y' לקבלת נתונים שנתיים."
                )
            else:
                st.warning("אין נתונים לסינון הנוכחי.")
            return

        # Quick stats row
        n_dates  = display_df["date"].nunique()
        n_yearly = (display_df["frequency"] == "yearly").sum()  if "frequency" in display_df.columns else 0
        n_monthly = (display_df["frequency"] == "monthly").sum() if "frequency" in display_df.columns else 0
        sc1, sc2, sc3 = st.columns(3)
        sc1.metric("נקודות זמן", n_dates)
        sc2.metric("נתונים חודשיים", n_monthly // max(1, display_df["allocation_name"].nunique()))
        sc3.metric("נתונים שנתיים",  n_yearly  // max(1, display_df["allocation_name"].nunique()))

        # ── Tabs ──────────────────────────────────────────────────────────
        t_ts, t_snap, t_delta, t_heat, t_stats, t_rank = st.tabs([
            "📈 סדרת זמן",
            "📍 Snapshot",
            "🔄 שינוי / Delta",
            "🌡️ Heatmap",
            "📊 סטטיסטיקות",
            "🏆 דירוג",
        ])

        # ── Tab 1: Time series ────────────────────────────────────────────
        with t_ts:
            from institutional_strategy_analysis.charts import build_timeseries
            fig = build_timeseries(display_df)
            _safe_plotly(fig, key="isa_ts")
            st.caption(
                "קווים מלאים = נתונים חודשיים | קווים מקווקוים = נתונים שנתיים. "
                "שנים שמכוסות על ידי נתונים חודשיים לא מוצגות כשנתיות."
            )
            col_dl, _ = st.columns([1, 5])
            with col_dl:
                st.download_button("⬇️ CSV", data=_csv(display_df),
                                   file_name="isa_timeseries.csv", mime="text/csv",
                                   key="isa_dl_ts")

        # ── Tab 2: Snapshot ───────────────────────────────────────────────
        with t_snap:
            max_d = display_df["date"].max().date()
            min_d = display_df["date"].min().date()
            snap_date = st.date_input(
                "תאריך Snapshot",
                value=max_d, min_value=min_d, max_value=max_d,
                help="מציג את הערך האחרון הידוע עד לתאריך שנבחר.",
                key="isa_snap_date",
            )
            from institutional_strategy_analysis.charts import build_snapshot
            _safe_plotly(build_snapshot(display_df, pd.Timestamp(snap_date)), key="isa_snap")

            snap_df = display_df[display_df["date"] <= pd.Timestamp(snap_date)]
            if not snap_df.empty:
                i = snap_df.groupby(["manager", "track", "allocation_name"])["date"].idxmax()
                tbl = snap_df.loc[i][["manager", "track", "allocation_name",
                                       "allocation_value", "date"]].copy()
                tbl["date"] = tbl["date"].dt.strftime("%Y-%m")
                tbl.columns = ["מנהל", "מסלול", "רכיב", "ערך (%)", "תאריך"]
                st.dataframe(tbl.sort_values("ערך (%)", ascending=False)
                               .reset_index(drop=True),
                             use_container_width=True, hide_index=True)

        # ── Tab 3: Delta ──────────────────────────────────────────────────
        with t_delta:
            min_d = display_df["date"].min().date()
            max_d = display_df["date"].max().date()
            dc1, dc2 = st.columns(2)
            with dc1:
                date_a = st.date_input("תאריך A (מוצא)",
                                       value=_clamp(max_d - timedelta(days=365), min_d, max_d),
                                       min_value=min_d, max_value=max_d,
                                       help="תאריך ההתחלה להשוואה.",
                                       key="isa_da")
            with dc2:
                date_b = st.date_input("תאריך B (יעד)", value=max_d,
                                       min_value=min_d, max_value=max_d,
                                       help="תאריך הסיום להשוואה.",
                                       key="isa_db")
            if date_a >= date_b:
                st.warning("תאריך A חייב להיות לפני B.")
            else:
                from institutional_strategy_analysis.charts import build_delta
                fig_d, delta_tbl = build_delta(display_df,
                                                pd.Timestamp(date_a),
                                                pd.Timestamp(date_b))
                _safe_plotly(fig_d, key="isa_delta")
                if not delta_tbl.empty:
                    st.dataframe(delta_tbl.reset_index(drop=True),
                                 use_container_width=True, hide_index=True)
                    col_dl2, _ = st.columns([1, 5])
                    with col_dl2:
                        st.download_button("⬇️ CSV", data=_csv(delta_tbl),
                                           file_name="isa_delta.csv", mime="text/csv",
                                           key="isa_dl_delta")

        # ── Tab 4: Heatmap ────────────────────────────────────────────────
        with t_heat:
            from institutional_strategy_analysis.charts import build_heatmap
            heat_df = display_df.copy()
            if display_df["date"].nunique() > 48:
                cutoff = display_df["date"].max() - pd.DateOffset(months=48)
                heat_df = display_df[display_df["date"] >= cutoff]
                st.caption("מוצגים 48 חודשים אחרונים. בחר 'הכל' לצפייה מלאה.")
            _safe_plotly(build_heatmap(heat_df), key="isa_heat")

        # ── Tab 5: Summary stats ──────────────────────────────────────────
        with t_stats:
            from institutional_strategy_analysis.charts import build_summary_stats
            stats = build_summary_stats(display_df)
            if stats.empty:
                st.info("אין מספיק נתונים לסטטיסטיקה.")
            else:
                st.dataframe(stats.reset_index(drop=True),
                             use_container_width=True, hide_index=True)
                col_dl3, _ = st.columns([1, 5])
                with col_dl3:
                    st.download_button("⬇️ CSV", data=_csv(stats),
                                       file_name="isa_stats.csv", mime="text/csv",
                                       key="isa_dl_stats")

        # ── Tab 6: Ranking ────────────────────────────────────────────────
        with t_rank:
            from institutional_strategy_analysis.charts import build_ranking
            if display_df["allocation_name"].nunique() > 1:
                rank_alloc = st.selectbox(
                    "רכיב לדירוג",
                    options=sorted(display_df["allocation_name"].unique()),
                    help="בחר רכיב שלפיו יוצג הדירוג החודשי.",
                    key="isa_rank_alloc",
                )
                rank_df = display_df[display_df["allocation_name"] == rank_alloc]
            else:
                rank_df = display_df

            _safe_plotly(
                build_ranking(rank_df,
                              title=f"דירוג מנהלים — {rank_df['allocation_name'].iloc[0]}"
                              if not rank_df.empty else "דירוג"),
                key="isa_rank",
            )

            # Volatility table
            if not rank_df.empty:
                vol = []
                for (mgr, trk), g in rank_df.groupby(["manager", "track"]):
                    chg = g.sort_values("date")["allocation_value"].diff().dropna()
                    vol.append({
                        "מנהל": mgr, "מסלול": trk,
                        "תנודתיות (STD)": round(chg.std(), 3) if len(chg) > 1 else float("nan"),
                        "שינוי מקסימלי": round(chg.abs().max(), 3) if not chg.empty else float("nan"),
                    })
                if vol:
                    st.caption("תנודתיות לפי מנהל:")
                    st.dataframe(
                        pd.DataFrame(vol).sort_values("תנודתיות (STD)", ascending=False)
                          .reset_index(drop=True),
                        use_container_width=True, hide_index=True,
                    )

        # ── Raw data ──────────────────────────────────────────────────────
        with st.expander("📋 נתונים גולמיים", expanded=False):
            disp = display_df.copy()
            if "date" in disp.columns:
                disp["date"] = disp["date"].dt.strftime("%Y-%m-%d")
            st.dataframe(disp.reset_index(drop=True),
                         use_container_width=True, hide_index=True)
            st.download_button("⬇️ ייצוא כל הנתונים", data=_csv(display_df),
                               file_name="isa_all.csv", mime="text/csv",
                               key="isa_dl_all")

        # ── AI Analysis (full dataset for peer comparison) ────────────────
        # Build full_df — all managers, same track(s) — for relative peer analysis
        all_filters = {
            "managers":        opts["managers"],
            "tracks":          sel_tracks,
            "allocation_names": sel_allocs,
        }
        try:
            full_df = _build_series(df_yearly, df_monthly, sel_range, custom_start, all_filters)
        except Exception:
            full_df = display_df

        ai_context = {
            "managers":         sel_mgr,
            "tracks":           sel_tracks,
            "allocation_names": sel_allocs,
            "selected_range":   sel_range,
            "date_min":         display_df["date"].min().strftime("%Y-%m") if not display_df.empty else "",
            "date_max":         display_df["date"].max().strftime("%Y-%m") if not display_df.empty else "",
        }
        _render_ai_section(display_df, full_df, ai_context, sel_mgr, sel_tracks)
