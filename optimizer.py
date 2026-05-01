import math
import pandas as pd

ALL_COMBOS = ['JH', 'JH+PV', 'Gyal', 'Gyal+PV', 'Senec', 'Senec+PV', 'Senec+PV+Prešov', 'Senec+Prešov']

DEFAULT_COMBOS = set(ALL_COMBOS)


def _nveh(pallets: int, cap: int) -> int:
    return math.ceil(pallets / cap) if pallets > 0 else 0


def _assign_pv(jh: int, gyal: int, senec: int, pv: int, allowed_targets: set, cap: int) -> str | None:
    """Return which main depot PV should join to minimise total vehicle count."""
    if pv == 0 or not allowed_targets:
        return None
    best, best_score = None, (9999, 9999)
    for name, main in [('JH', jh), ('Gyal', gyal), ('Senec', senec)]:
        if name not in allowed_targets or main == 0:
            continue
        others = [v for k, v in [('JH', jh), ('Gyal', gyal), ('Senec', senec)] if k != name]
        total_v = _nveh(main + pv, cap) + sum(_nveh(o, cap) for o in others)
        waste = total_v * cap - (jh + gyal + senec + pv)
        if (total_v, waste) < best_score:
            best_score = (total_v, waste)
            best = name
    return best


def _scale_bd(bd: dict, frac: float, target: int) -> dict:
    items = list(bd.items())
    out, rem = {}, target
    for d, p in items[:-1]:
        v = max(0, min(round(p * frac), rem))
        out[d] = v
        rem -= v
    out[items[-1][0]] = max(0, rem)
    return out


def _bd_str(bd: dict) -> str:
    return ' | '.join(f'{d}: {p}p' for d, p in bd.items() if p > 0)


def optimise(df: pd.DataFrame, allowed_combos: set = DEFAULT_COMBOS, capacity: int = 33):
    """
    Returns (trips_df, warnings: list[str]).

    allowed_combos controls which route types are permitted.
    PV assignment only considers depots whose combo is enabled.
    Prešov is always routed via Senec when Senec+Prešov or Senec+PV+Prešov is active.
    """
    df = df.copy()
    df['datum dodání'] = pd.to_datetime(df['datum dodání'])

    # Derive PV routing targets from enabled combos
    pv_targets: set[str] = set()
    if 'JH+PV' in allowed_combos:
        pv_targets.add('JH')
    if 'Gyal+PV' in allowed_combos:
        pv_targets.add('Gyal')
    if 'Senec+PV' in allowed_combos or 'Senec+PV+Prešov' in allowed_combos:
        pv_targets.add('Senec')

    presov_ok = 'Senec+Prešov' in allowed_combos or 'Senec+PV+Prešov' in allowed_combos

    agg = (
        df.groupby(['datum dodání', 'DEPO'])
        .agg(palety=('palety', 'sum'), n_zas=('palety', 'count'))
        .reset_index()
    )

    rows: list[dict] = []
    warnings: list[str] = []

    for date, day in agg.groupby('datum dodání'):
        dp = day.set_index('DEPO')
        p = lambda d: int(dp.loc[d, 'palety']) if d in dp.index else 0  # noqa: E731
        z = lambda d: int(dp.loc[d, 'n_zas'])   if d in dp.index else 0  # noqa: E731

        jh, pv, gyal, senec, presov = p('JH'), p('PV'), p('Gyal'), p('Senec'), p('Prešov')
        zj, zp, zg, zs, zpr = z('JH'), z('PV'), z('Gyal'), z('Senec'), z('Prešov')

        # Prešov handling
        if presov > 0 and not presov_ok:
            warnings.append(
                f'{date.date()}: {presov} palet Prešov nelze přepravit '
                f'(žádná aktivní Senec+Prešov kombinace)'
            )

        senec_base = senec + (presov if presov_ok else 0)

        # PV assignment
        pv_target = _assign_pv(jh, gyal, senec_base, pv, pv_targets, capacity) if pv > 0 else None
        if pv > 0 and pv_target is None and pv_targets:
            # pv_targets exist but no main depot has pallets on this day
            warnings.append(
                f'{date.date()}: {pv} palet PV nelze přiřadit '
                f'(partnerské depo má 0 palet, nebo žádná PV kombinace není aktivní)'
            )
        elif pv > 0 and not pv_targets:
            warnings.append(
                f'{date.date()}: {pv} palet PV nelze přepravit (žádná PV kombinace není aktivní)'
            )

        # ── Build groups ──────────────────────────────────────────────────────

        def add_trips(combo: str, total_p: int, bd: dict, total_z: int):
            if total_p <= 0 or combo is None:
                return
            nv = _nveh(total_p, capacity)
            base, extra = total_p // nv, total_p % nv
            for v in range(nv):
                vp = base + (1 if v < extra else 0)
                frac = vp / total_p
                vbd = _scale_bd(bd, frac, vp)
                v_z = round(total_z * frac) if nv > 1 else total_z
                rows.append({
                    'Datum': date,
                    'Měsíc': date.month,
                    'Kombinace dep': combo,
                    'Počet dep': sum(1 for x in bd.values() if x > 0),
                    'Palety': vp,
                    'Vytížení': round(vp / capacity, 6),
                    'Počet zásilek': max(1, v_z),
                    'Rozpis palet po depech': _bd_str(vbd),
                    'Pozn.': '',
                })

        # JH group
        pv_jh = pv_target == 'JH'
        if jh > 0 or pv_jh:
            if pv_jh and jh > 0:
                add_trips('JH+PV', jh + pv, {'JH': jh, 'PV': pv}, zj + zp)
            elif pv_jh:
                pass  # PV can't go to JH alone (no JH pallets) – already warned
            elif jh > 0:
                if 'JH' in allowed_combos:
                    add_trips('JH', jh, {'JH': jh}, zj)
                else:
                    warnings.append(f'{date.date()}: {jh} palet JH nelze přepravit (kombinace JH není aktivní)')

        # Gyal group
        pv_gyal = pv_target == 'Gyal'
        if gyal > 0 or pv_gyal:
            if pv_gyal and gyal > 0:
                add_trips('Gyal+PV', gyal + pv, {'Gyal': gyal, 'PV': pv}, zg + zp)
            elif pv_gyal:
                pass
            elif gyal > 0:
                if 'Gyal' in allowed_combos:
                    add_trips('Gyal', gyal, {'Gyal': gyal}, zg)
                else:
                    warnings.append(f'{date.date()}: {gyal} palet Gyal nelze přepravit (kombinace Gyal není aktivní)')

        # Senec group
        pv_senec = pv_target == 'Senec'
        has_presov = presov > 0 and presov_ok

        senec_parts: dict[str, int] = {}
        if senec > 0:
            senec_parts['Senec'] = senec
        if has_presov:
            senec_parts['Prešov'] = presov
        if pv_senec:
            senec_parts['PV'] = pv

        if senec_parts:
            keys = frozenset(senec_parts)
            if keys == frozenset({'Senec', 'PV', 'Prešov'}):
                combo = 'Senec+PV+Prešov'
            elif keys == frozenset({'Senec', 'PV'}):
                combo = 'Senec+PV'
            elif keys == frozenset({'Senec', 'Prešov'}):
                combo = 'Senec+Prešov'
            elif keys == frozenset({'Senec'}):
                combo = 'Senec' if 'Senec' in allowed_combos else None
                if combo is None:
                    warnings.append(f'{date.date()}: {senec} palet Senec nelze přepravit (kombinace Senec není aktivní)')
            elif keys == frozenset({'Prešov'}):
                combo = 'Senec+Prešov'  # only Prešov, no Senec pallets
            else:
                combo = None

            z_senec = zs + (zpr if has_presov else 0) + (zp if pv_senec else 0)
            total_senec = sum(senec_parts.values())
            add_trips(combo, total_senec, senec_parts, z_senec)

    trips = pd.DataFrame(rows) if rows else pd.DataFrame(
        columns=['Datum', 'Měsíc', 'Kombinace dep', 'Počet dep',
                 'Palety', 'Vytížení', 'Počet zásilek', 'Rozpis palet po depech', 'Pozn.']
    )
    if not trips.empty:
        trips.insert(0, 'Č. jízdy', range(1, len(trips) + 1))

    return trips, warnings


def make_matrix(trips: pd.DataFrame) -> pd.DataFrame:
    if trips.empty:
        return pd.DataFrame()
    pivot = trips.pivot_table(
        index='Kombinace dep', columns='Měsíc',
        values='Palety', aggfunc='sum', fill_value=0
    )
    pivot.columns = [int(c) for c in pivot.columns]
    # Ensure all 12 months present
    for m in range(1, 13):
        if m not in pivot.columns:
            pivot[m] = 0
    pivot = pivot[[m for m in range(1, 13)]]
    pivot['Celkem'] = pivot.sum(axis=1)
    pivot = pivot.sort_values('Celkem', ascending=False)
    pivot.index.name = 'Kombinace'
    return pivot.reset_index()
