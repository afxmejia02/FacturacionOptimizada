"""Cruce de desprendibles contra transferencias y seguridad social."""
from __future__ import annotations

from collections import Counter
import pandas as pd
import re

from .formato import _limpiar_doc
from .depuracion import log as _log


def conciliar(df_desprendibles, df_transferencia, df_seguridad=None):
    """
    Concilia los datos de desprendibles y transferencias e incorpora Devengado e IBC.
    """
    def _normalizar_numero(valor):
        if valor is None:
            return None
        try:
            if pd.isna(valor):
                return None
        except Exception:
            pass

        try:
            numero = float(valor)
        except Exception:
            return None

        if numero.is_integer():
            return int(numero)
        return numero

    def _normalizar_lista(valores):
        normalizados = []
        vistos = set()
        for valor in valores or []:
            numero = _normalizar_numero(valor)
            if numero is None or numero in vistos:
                continue
            vistos.add(numero)
            normalizados.append(numero)
        return normalizados

    def _normalizar_lista_completa(valores):
        """Como ``_normalizar_lista`` pero **sin deduplicar**.

        Para seguridad social los devengados/IBC pueden repetirse entre
        quincenas (dos quincenas con el mismo valor) y deben contarse ambas.
        """
        return [
            numero
            for numero in (_normalizar_numero(v) for v in (valores or []))
            if numero is not None
        ]

    resultados = []
    resultados_transfers = []
    resultados_seguridad = []

    # Limpiar números de documento
    if df_desprendibles is None or df_desprendibles.empty:
        return pd.DataFrame(resultados)

    def _clave_cuenta(valor):
        """Clave de cuenta robusta: solo dígitos y sin ceros a la izquierda.

        Permite cruzar la cuenta del desprendible contra el número de producto
        del soporte aunque difieran en ceros de relleno o en formato.
        """
        if valor is None:
            return ""
        try:
            if pd.isna(valor):
                return ""
        except (TypeError, ValueError):
            pass
        return re.sub(r"\D", "", str(valor)).lstrip("0")

    df_desprendibles["Identificacion"] = _limpiar_doc(df_desprendibles["Identificacion"])

    hay_transferencias = df_transferencia is not None and not df_transferencia.empty
    if hay_transferencias:
        df_transferencia["Documento"] = _limpiar_doc(df_transferencia["Documento"])
        # Clave de cuenta normalizada para el cruce por producto/cuenta.
        df_transferencia["_cuenta_key"] = (
            df_transferencia["Cuenta"].map(_clave_cuenta)
            if "Cuenta" in df_transferencia.columns
            else ""
        )
        _log(
            f"[reconcile] {len(df_transferencia)} transferencias disponibles; "
            f"{df_transferencia['Documento'].nunique()} documento(s) distinto(s)."
        )
    else:
        _log("[reconcile] No se recibieron transferencias para cruzar.")

    if df_seguridad is not None and not df_seguridad.empty:
        df_seguridad["cc"] = (
            df_seguridad["cc"].astype(str).str.replace(r"[^\d]", "", regex=True).str.lstrip("0")
        )

    # Agrupar desprendibles por identificación y conciliar con transferencias
    for doc, grupo_despr in df_desprendibles.groupby("Identificacion"):
        netos = _normalizar_lista(grupo_despr["Neto"].dropna().tolist())
        suma_netos = _normalizar_numero(sum(netos)) if netos else None
        # Devengado: sumar TODOS los valores SIN deduplicar. Dos quincenas con
        # el mismo devengado deben contarse ambas (a diferencia del cruce de
        # transferencias, donde sí se deduplica).
        devs = _normalizar_lista_completa(grupo_despr["Devengado"].dropna().tolist()) if "Devengado" in grupo_despr else []
        sum_devs = _normalizar_numero(sum(devs)) if devs else None

        cta = grupo_despr["Cuenta"].iloc[0] if "Cuenta" in grupo_despr.columns else None
        cta_key = _clave_cuenta(cta)

        # Ventana de periodo de esta persona, tomada de SUS desprendibles
        # ("Periodo: inicio al fin"). Sirve para descartar transferencias de
        # otras quincenas/meses que comparten el mismo documento.
        win_ini = win_fin = None
        if "PeriodoInicio" in grupo_despr.columns and "PeriodoFin" in grupo_despr.columns:
            inis = pd.to_datetime(grupo_despr["PeriodoInicio"], errors="coerce").dropna()
            fins = pd.to_datetime(grupo_despr["PeriodoFin"], errors="coerce").dropna()
            if not inis.empty and not fins.empty:
                win_ini, win_fin = inis.min(), fins.max()

        # Buscar transferencias coincidentes (por documento o por cuenta) y
        # comparar por suma. El cruce por documento es el principal; el de
        # cuenta es un respaldo para soportes donde el documento difiera.
        grupo_trans = pd.DataFrame()
        estado_trans = "Transferencia no encontrada"
        valores_trans = []
        suma_trans = None
        if hay_transferencias:
            por_documento = df_transferencia["Documento"] == doc
            por_cuenta = (
                (df_transferencia["_cuenta_key"] == cta_key) & (cta_key != "")
                if "_cuenta_key" in df_transferencia.columns
                else False
            )
            grupo_trans = df_transferencia[por_documento | por_cuenta]

        # Separar transferencias confiables (con etiqueta NÓMINA) de las
        # candidatas (otro layout sin etiqueta ni fecha-factura de quincena).
        # Si no hay columna 'EsNomina' (formato TABARCA o datos antiguos), se
        # tratan todas como confiables -> comportamiento previo intacto.
        if not grupo_trans.empty and "EsNomina" in grupo_trans.columns:
            es_nom = grupo_trans["EsNomina"].fillna(True).astype(bool)
            trans_confiables = grupo_trans[es_nom]
            trans_candidatas = grupo_trans[~es_nom]
        else:
            trans_confiables = grupo_trans
            trans_candidatas = grupo_trans.iloc[0:0]

        # Filtrar por periodo SOLO las confiables: conservar las que caen dentro
        # de la ventana de los desprendibles (las sin fecha legible se conservan).
        # Evita sumar quincenas/meses ajenos al soporte.
        if (
            not trans_confiables.empty
            and win_ini is not None
            and win_fin is not None
            and "Fecha" in trans_confiables.columns
        ):
            fechas = pd.to_datetime(trans_confiables["Fecha"], errors="coerce")
            en_ventana = fechas.isna() | ((fechas >= win_ini) & (fechas <= win_fin))
            descartadas = int((~en_ventana).sum())
            if descartadas:
                _log(
                    f"[reconcile] doc={doc}: {descartadas} transferencia(s) fuera del "
                    f"periodo [{win_ini.date()}..{win_fin.date()}] descartada(s)."
                )
            trans_confiables = trans_confiables[en_ventana]

        # Candidatas: solo valen si su valor coincide con un neto del desprendible
        # que aún no haya sido cubierto por una transferencia confiable. Así se
        # rescata el pago real (mismo valor que el neto) aunque el renglón no
        # traiga etiqueta/fecha, y se descartan importes de otra quincena.
        candidatas_validas = trans_candidatas.iloc[0:0]
        if not trans_candidatas.empty:
            val_conf = _normalizar_lista(trans_confiables["Valor"].dropna().tolist()) if not trans_confiables.empty else []
            pendientes = Counter(int(round(float(n))) for n in netos if pd.notna(n))
            pendientes.subtract(Counter(int(round(float(v))) for v in val_conf))
            idx_keep = []
            for idx_c, val_c in trans_candidatas["Valor"].dropna().items():
                clave = int(round(float(val_c)))
                if pendientes.get(clave, 0) > 0:
                    idx_keep.append(idx_c)
                    pendientes[clave] -= 1
                else:
                    _log(
                        f"[reconcile] doc={doc}: transferencia candidata {val_c} sin neto "
                        f"que la respalde -> descartada (posible otra quincena)."
                    )
            candidatas_validas = trans_candidatas.loc[idx_keep]

        grupo_trans = pd.concat([trans_confiables, candidatas_validas])

        if not grupo_trans.empty:
            valores_trans = _normalizar_lista(grupo_trans["Valor"].dropna().tolist())
            # Sumar valores y comparar con la suma de netos.
            try:
                suma_trans = _normalizar_numero(sum(valores_trans)) if valores_trans else None
            except Exception:
                suma_trans = None

            if suma_netos is not None and suma_trans is not None and suma_netos == suma_trans:
                estado_trans = "OK"
            else:
                estado_trans = "Valor no coincide"
                _log(
                    f"[reconcile] doc={doc}: transferencia encontrada pero la suma no "
                    f"coincide (netos={suma_netos} vs transferencias={suma_trans})."
                )
        else:
            valores_trans = None
            if hay_transferencias:
                # Log diagnóstico: la transferencia existe en el universo pero no
                # cruzó por documento ni por cuenta para esta persona.
                _log(
                    f"[reconcile] doc={doc} (cuenta={cta_key or '—'}): sin transferencia "
                    f"que cruce por documento ni por cuenta -> 'Transferencia no encontrada'."
                )

        # Obtener IBCs desde df_seguridad si está disponible. NO se deduplican:
        # se suman todos (igual que el cruce de transferencias).
        ibc_vals = []
        if df_seguridad is not None and not df_seguridad.empty:
            matches = df_seguridad[df_seguridad["cc"] == doc]
            if not matches.empty:
                ibc_vals = _normalizar_lista_completa(matches["ibc"].dropna().tolist())
        suma_ibc = _normalizar_numero(sum(ibc_vals)) if ibc_vals else None

        # Estado seguridad social: comparar SUMA de devengados vs SUMA de IBC.
        if sum_devs is None:
            estado_seg = "Devengado no encontrado"
        elif suma_ibc is not None and sum_devs == suma_ibc:
            estado_seg = "OK"
        else:
            estado_seg = "Devengado no coincide"

        # Diferencia = devengado(desprendible) - IBC(seguridad). Lo ausente cuenta como 0.
        diferencia_seg = (
            _normalizar_numero((sum_devs or 0) - (suma_ibc or 0))
            if sum_devs is not None else None
        )
        # Diferencia = neto(desprendible) - transferencia. Lo ausente cuenta como 0.
        diferencia_trans = (
            _normalizar_numero((suma_netos or 0) - (suma_trans or 0))
            if (suma_netos is not None or suma_trans is not None) else None
        )

        # Agregar resultado para transferencias (usa suma para OK)
        resultados_transfers.append({
            "Identificación": doc,
            "Cuenta": cta,
            "Estado": estado_trans,
            "Neto_desprendibles": list(netos),
            "Valores_transferencia": list(valores_trans) if valores_trans is not None else None,
            "Diferencia": diferencia_trans,
        })

        # Agregar resultado para seguridad social (IBC)
        resultados_seguridad.append({
            "Identificación": doc,
            "Estado": estado_seg,
            "Devengado": sum_devs,
            "IBC": ibc_vals if ibc_vals else None,
            "Diferencia": diferencia_seg,
        })
    df_t = pd.DataFrame(resultados_transfers)
    df_s = pd.DataFrame(resultados_seguridad)

    return df_t, df_s
