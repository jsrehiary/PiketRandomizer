import pandas as pd
from datetime import datetime, timedelta
import random

# =====================
# KONFIGURASI
# =====================
START_DATE = datetime(2026, 1, 5)
PERIOD_DAYS = 14
TOTAL_PERIODS = 12
RANDOM_SEED = 42

random.seed(RANDOM_SEED)

# =====================
# LOAD CSV
# =====================
df = pd.read_csv("members.csv")

# semua anggota unik (buat kolom checklist)
all_members = df["nama"].unique().tolist()

# mapping divisi -> anggota
divisi_map = (
    df.groupby("divisi")["nama"]
      .apply(list)
      .to_dict()
)

# =====================
# RANDOMIZE PER DIVISI
# =====================
for divisi in divisi_map:
    random.shuffle(divisi_map[divisi])

# =====================
# GENERATE JADWAL
# =====================
schedule = []
checklist_rows = []

for period in range(TOTAL_PERIODS):
    start = START_DATE + timedelta(days=period * PERIOD_DAYS)
    end = start + timedelta(days=PERIOD_DAYS - 1)

    for divisi, anggota in divisi_map.items():
        petugas = anggota[period % len(anggota)]

        # sheet jadwal
        schedule.append({
            "Periode": period + 1,
            "Mulai": start.strftime("%Y-%m-%d"),
            "Selesai": end.strftime("%Y-%m-%d"),
            "Divisi": divisi,
            "Petugas": petugas
        })

        # sheet checklist
        row = {
            "Periode": period + 1,
            "Divisi": divisi
        }

        for member in all_members:
            row[member] = "✔" if member == petugas else ""

        checklist_rows.append(row)

# =====================
# DATAFRAME
# =====================
jadwal_df = pd.DataFrame(schedule)
checklist_df = pd.DataFrame(checklist_rows)

# =====================
# EXPORT KE EXCEL
# =====================
with pd.ExcelWriter(
    "jadwal_piket.xlsx",
    engine="openpyxl"
) as writer:
    jadwal_df.to_excel(writer, sheet_name="Jadwal_Piket", index=False)
    checklist_df.to_excel(writer, sheet_name="Checklist", index=False)

print("Excel berhasil dibuat: jadwal_piket.xlsx")