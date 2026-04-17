# -*- coding: utf-8 -*-
"""JSON persistence for students, weeks, plans, notes and topics."""

import json
import os
import shutil
from datetime import datetime
from typing import Any, Dict, Optional

from constants import GUNLER, JSON_DOSYA

# Varsayılan konu verisi (JSON'da yoksa boş; uygulama ilk açılışta kendi listesini yazabilir)
DEFAULT_KONU_VERISI: Dict[str, Any] = {}

_TIME_ALIASES = {
    "8:00": "08:00",
    "9:00": "09:00",
}

_DAY_ALIASES = {
    "Pazartesi": "Monday",
    "Salı": "Tuesday",
    "Çarşamba": "Wednesday",
    "Perşembe": "Thursday",
    "Cuma": "Friday",
    "Cumartesi": "Saturday",
    "Pazar": "Sunday",
}


def _normalize_week_label(label: Any) -> str:
    text = str(label).strip()
    lower = text.lower()
    if lower.startswith("hafta "):
        suffix = text.split(" ", 1)[1] if " " in text else ""
        return f"Week {suffix}".strip()
    return text


def _normalize_hour_label(label: Any) -> str:
    """Convert hour labels to HH:MM format (e.g. 9:00 -> 09:00)."""
    s = str(label)
    parca = s.split(":")
    if len(parca) == 2 and parca[0].isdigit() and parca[1].isdigit():
        return f"{int(parca[0]):02d}:{int(parca[1]):02d}"
    return _TIME_ALIASES.get(s, s)


def _normalize_entry(entry: Any) -> Dict[str, Any]:
    """Normalize entry to {text, done} format (legacy string compatible)."""
    if isinstance(entry, dict):
        return {"text": str(entry.get("text", "")), "done": bool(entry.get("done", False))}
    if entry is None:
        return {"text": "", "done": False}
    return {"text": str(entry), "done": False}


def _normalize_program(program: Dict[str, Any]) -> None:
    """Normalize one weekly plan as day/hour -> {text, done}."""
    for old_day, new_day in _DAY_ALIASES.items():
        if old_day in program and new_day not in program and isinstance(program.get(old_day), dict):
            program[new_day] = program.get(old_day, {})
        if old_day in program:
            del program[old_day]
    for gun in GUNLER:
        gun_prog = program.get(gun)
        if gun_prog is None:
            program[gun] = {}
            gun_prog = program[gun]
        elif not isinstance(gun_prog, dict):
            program[gun] = {}
            gun_prog = program[gun]
        for saat, deger in list(gun_prog.items()):
            yeni_saat = _normalize_hour_label(saat)
            entry = _normalize_entry(deger)
            if yeni_saat != saat:
                del gun_prog[saat]
                if yeni_saat not in gun_prog:
                    gun_prog[yeni_saat] = entry
            else:
                gun_prog[saat] = entry


def _normalize_tum_programlar(ogrenciler: Dict[str, Any]) -> None:
    """Normalize all student plans and migrate legacy structures."""
    for ad, val in list(ogrenciler.items()):
        gun_anahtarlari = [k for k in val.keys() if k in GUNLER]
        if gun_anahtarlari:
            # Legacy format: root-level day keys -> nest under Week 1
            eski_program = {g: val.get(g, {}) for g in GUNLER}
            _normalize_program(eski_program)
            ogrenciler[ad] = {"Week 1": eski_program}
        else:
            normalized_weeks: Dict[str, Any] = {}
            for hafta_adi, program in val.items():
                normalized_name = _normalize_week_label(hafta_adi)
                if isinstance(program, dict):
                    _normalize_program(program)
                    normalized_weeks[normalized_name] = program
            ogrenciler[ad] = normalized_weeks


class JsonStorage:
    """Ana veri dosyasının okunması ve yazılması."""

    def __init__(self, base_dir: Optional[str] = None):
        self.base_dir = base_dir or os.getcwd()
        self._path = os.path.join(self.base_dir, JSON_DOSYA)
        self._backup_dir = os.path.join(self.base_dir, "backups")

    @staticmethod
    def _default_payload() -> Dict[str, Any]:
        return {
            "ogrenciler": {},
            "ogrenci_notlari": {},
            "ogrenci_alanlari": {},
            "ogrenci_tercihleri": {},
            "denemeler": {},
            "aktif_ogrenci": None,
            "aktif_hafta": None,
            "konu_verisi": DEFAULT_KONU_VERISI,
        }

    def path_exists(self) -> bool:
        return os.path.exists(self._path)

    def load(self) -> Dict[str, Any]:
        """Veriyi yükle. Dosya yoksa veya bozuksa varsayılan yapı döner."""
        out = self._default_payload()
        if not self.path_exists():
            return out
        try:
            with open(self._path, "r", encoding="utf-8") as f:
                data = json.load(f)
        except Exception:
            return out
        out["ogrenciler"] = data.get("ogrenciler", {})
        out["ogrenci_notlari"] = data.get("ogrenci_notlari", {})
        out["ogrenci_alanlari"] = data.get("ogrenci_alanlari", {})
        out["ogrenci_tercihleri"] = data.get("ogrenci_tercihleri", {})
        out["denemeler"] = data.get("denemeler", {})
        out["aktif_ogrenci"] = data.get("aktif_ogrenci")
        out["aktif_hafta"] = data.get("aktif_hafta")
        out["konu_verisi"] = data.get("konu_verisi") or DEFAULT_KONU_VERISI
        _normalize_tum_programlar(out["ogrenciler"])
        if out["ogrenciler"] and out["aktif_ogrenci"] not in out["ogrenciler"]:
            out["aktif_ogrenci"] = next(iter(out["ogrenciler"].keys()))
        return out

    def save(self, data: Dict[str, Any]) -> None:
        """Veriyi kaydet. Hata durumunda exception fırlatır."""
        payload = {
            "ogrenciler": data["ogrenciler"],
            "ogrenci_notlari": data.get("ogrenci_notlari", {}),
            "ogrenci_alanlari": data.get("ogrenci_alanlari", {}),
            "ogrenci_tercihleri": data.get("ogrenci_tercihleri", {}),
            "denemeler": data.get("denemeler", {}),
            "konu_verisi": data.get("konu_verisi", {}),
            "aktif_ogrenci": data.get("aktif_ogrenci"),
            "aktif_hafta": data.get("aktif_hafta"),
        }
        with open(self._path, "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)

    def _read_main_text(self) -> Optional[str]:
        if not self.path_exists():
            return None
        try:
            with open(self._path, "r", encoding="utf-8") as f:
                return f.read()
        except Exception:
            return None

    def backup(self) -> None:
        """Yedek oluşturur: hem son yedek dosyası hem tarihli yedek."""
        data = self._read_main_text()
        if data is None:
            return
        # Eski tek-dosya yedek (geri uyum)
        legacy_backup_path = os.path.join(self.base_dir, "program_kayitlari_backup.json")
        try:
            with open(legacy_backup_path, "w", encoding="utf-8") as f:
                f.write(data)
        except Exception:
            pass
        # Yeni tarihli yedekler
        try:
            os.makedirs(self._backup_dir, exist_ok=True)
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            stamp_path = os.path.join(self._backup_dir, f"program_kayitlari_{ts}.json")
            with open(stamp_path, "w", encoding="utf-8") as f:
                f.write(data)
        except Exception:
            pass

    def export_to(self, target_path: str) -> None:
        """Ana json verisini verilen dosya yoluna export eder."""
        data = self._read_main_text()
        if data is None:
            data = json.dumps(self._default_payload(), ensure_ascii=False, indent=2)
        with open(target_path, "w", encoding="utf-8") as f:
            f.write(data)

    def import_from(self, source_path: str) -> None:
        """Verilen json dosyasını ana veri dosyası olarak içe aktarır."""
        with open(source_path, "r", encoding="utf-8") as f:
            data = json.load(f)
        # Temel yapıya normalize et
        payload = self._default_payload()
        payload["ogrenciler"] = data.get("ogrenciler", {})
        payload["ogrenci_notlari"] = data.get("ogrenci_notlari", {})
        payload["ogrenci_alanlari"] = data.get("ogrenci_alanlari", {})
        payload["ogrenci_tercihleri"] = data.get("ogrenci_tercihleri", {})
        payload["denemeler"] = data.get("denemeler", {})
        payload["aktif_ogrenci"] = data.get("aktif_ogrenci")
        payload["aktif_hafta"] = data.get("aktif_hafta")
        payload["konu_verisi"] = data.get("konu_verisi") or DEFAULT_KONU_VERISI
        _normalize_tum_programlar(payload["ogrenciler"])
        self.save(payload)

    def list_backups(self) -> list[str]:
        """Tarihli yedek dosyalarının tam yollarını (yeniden eskiye) döndürür."""
        if not os.path.isdir(self._backup_dir):
            return []
        entries = []
        for name in os.listdir(self._backup_dir):
            if name.startswith("program_kayitlari_") and name.endswith(".json"):
                entries.append(os.path.join(self._backup_dir, name))
        entries.sort(reverse=True)
        return entries

    def restore_backup(self, backup_path: str) -> None:
        """Seçilen yedek dosyasını ana veri dosyası olarak geri yükler."""
        if not os.path.exists(backup_path):
            raise FileNotFoundError(backup_path)
        shutil.copyfile(backup_path, self._path)

    def reset_data(self) -> None:
        """Tüm veriyi sıfırlar ve boş varsayılan yapıyı kaydeder."""
        self.save(self._default_payload())
