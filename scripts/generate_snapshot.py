import sys
from pathlib import Path

# Add project root to sys.path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import notice_lib as nl

sys.stdout.reconfigure(encoding='utf-8')

# Ensure tests/golden exists
golden_dir = Path("tests/golden")
golden_dir.mkdir(parents=True, exist_ok=True)

csv_path = golden_dir / "synthetic_2_records.csv"
csv_content = """SERIAL_NUMBER,E_CHALLAN_NUMBER,VEHICLE_NUMBER,VIOLATION_DATE,NAME,ADDRESS,MOBILE_NUMBER,FINE_AMOUNT,SECTION
8787,CG5010992607181306,CG07A1626,18/07/2026,RAKESH VISHVAKARMA,"HNO 68 RAUT PARA PS SAHASPUR LOHARA DIST KAWARDHA KAWARDHA CG 491995",9907420371,500,129/194(D)Driving without helmet
8788,CG5011002607021126,CG07A3523,02/07/2026,JITENDRA KUMAR,"HOUSE NO 492 WARD 13 NEAR OLD POLICE CHOWKI MAIN ROAD CIVIL LINES BILASPUR CHHATTISGARH 495001",9826145461,500,129/194(D)Driving without helmet
"""
csv_path.write_text(csv_content, encoding="utf-8-sig")

records = nl.read_records(csv_path)
print(f"Read {len(records)} records")

tmpl_path = Path("templates/Formate_with_Sign_SectionSplit.docx")
soffice = r"C:\Program Files\LibreOffice\program\soffice.exe"

run = nl.generate_all(
    input_path=csv_path,
    template_path=tmpl_path,
    output_dir=golden_dir,
    batch_size=1000,
    soffice_path=soffice,
    workers=1
)

print(f"Run completed in {run.total_seconds:.2f}s, Batches OK: {len(run.ok_batches)}")
for b in run.batches:
    print(f"Batch {b.batch_index}: ok={b.ok}, pages={b.actual_pages}, pdf={b.pdf_path}, error={b.error}")

if run.ok_batches:
    gen_pdf = Path(run.ok_batches[0].pdf_path)
    target_pdf = golden_dir / "snapshot.pdf"
    if target_pdf.exists():
        target_pdf.unlink()
    gen_pdf.rename(target_pdf)
    print(f"Renamed {gen_pdf.name} to {target_pdf.name}")
    print(f"Snapshot PDF size: {target_pdf.stat().st_size} bytes, pages: {nl.count_pdf_pages(target_pdf)}")
