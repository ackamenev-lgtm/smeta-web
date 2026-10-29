from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
COMMON = ROOT / "common"
SPR = COMMON / "Промт и тарификатор"
TARIF = SPR / "Тарификатор НЕОСТРОЙ v2 — рынок, бригады, нормо-часы.xlsx"
MATERIALS = SPR / "Материалы НЕОСТРОЙ — себестоимость и наценка.xlsx"
STAFF = SPR / "Расценки НЕОСТРОЙ — для сотрудников.xlsx"
PRICE_BASE = SPR / "База цен материалов.csv"
TITUL = SPR / "Титульник.docx"
REQUISITES = SPR / "Реквизиты - НЕОСТРОЙ ООО.docx"
OUTBOX = ROOT / "outbox"
CACHE = Path("/tmp/neostroi_cache")
