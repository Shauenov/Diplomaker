# Deployment Guide

Step-by-step instructions for deploying the Diploma Automation System in production environments.

---

## Pre-Deployment Checklist

### System Requirements

- **OS**: Windows 7+ or Linux (Ubuntu 18.04+) or macOS 10.12+
- **Python**: 3.8 - 3.11
- **RAM**: Minimum 4GB (8GB recommended for batch processing 1000+ students)
- **Disk Space**: 
  - Installation: ~150MB (Python + dependencies)
  - Per 100 diplomas: ~50MB (100 source + 200 output files)
- **Excel**: Not required (xlsxwriter handles generation)

### Pre-Installation

```bash
# 1. Verify Python installation
python --version
# Output: Python 3.8.x, 3.9.x, 3.10.x, or 3.11.x

# 2. Verify pip is available
pip --version
# Output: pip 20.0+

# 3. Create project directory
mkdir diploma_automation
cd diploma_automation

# 4. Verify write permissions
mkdir test_dir && rmdir test_dir  # Should succeed
```

---

## Installation Steps

### Step 1: Clone/Copy Project Files

```bash
# If using Git:
git clone <repository-url> .

# Or copy files manually:
# Copy all files from source to deployment directory
```

### Step 2: Create Python Virtual Environment

**Windows**:
```bash
python -m venv venv
venv\Scripts\activate
```

**Linux/macOS**:
```bash
python3 -m venv venv
source venv/bin/activate
```

### Step 3: Install Dependencies

```bash
pip install -r requirements.txt
```

**Expected Output**:
```
Successfully installed pandas-2.0.0 openpyxl-3.10.0 xlsxwriter-3.0.0 pyyaml-6.0
```

### Step 3.1 (Optional): Build Portable Desktop EXE

```powershell
powershell -ExecutionPolicy Bypass -File .\packaging\build_portable.ps1
```

Artifacts:
- `dist/DiplomaGenerator/` folder contains portable `DiplomaGenerator.exe` and bundled runtime files.
- Distribute the full folder contents (not only the exe).

### Step 4: Verify Installation

```bash
python -c "from config import settings, programs; print('✓ Configuration OK')"
python -c "from core import models, converters; print('✓ Core modules OK')"
```

**Expected Output**:
```
✓ Configuration OK
✓ Core modules OK
```

### Step 5: Configure Source File Path

Edit [config/settings.py](../config/settings.py):

```python
# Before (development):
# SOURCE_FILE = r"C:\Users\user\diploma_grades.xlsx"

# After (production):
SOURCE_FILE = r"\\server\shared\diplomas\2025-2026_grades.xlsx"
# OR (if local):
SOURCE_FILE = r"C:\Program Files\Diplomas\grades.xlsx"
```

### Step 6: Create Output Directory

```bash
# Create output directory
mkdir output_diplomas

# Verify permissions (should be writable)
echo "test" > output_diplomas/test.txt && del output_diplomas/test.txt
```

### Step 7: Test Configuration

```bash
python config/settings.py
```

**Expected Output**:
```
Source file: \\server\shared\diplomas\2025-2026_grades.xlsx
Output directory: C:\diploma_automation\output_diplomas
Grade thresholds: 10 entries loaded
Programs: 2 registered (IT, ACCOUNTING)
```

---

## Run First Generation

### Test with Sample Data

```bash
# Create test Excel file with 3 students
python -c "
import pandas as pd
data = {
    'No.': [1, 2, 3],
    'Full Name': ['Test Student 1', 'Test Student 2', 'Test Student 3'],
    'Қазақ тілі\nКазахский язык': [85, 92, 78],
}
df = pd.DataFrame(data)
df.to_excel('test_grades.xlsx', sheet_name='3F-1')
"

# Update config temporarily
# Set SOURCE_FILE = r"test_grades.xlsx"

# Generate diplomas
python -m batch._generate_it
```

**Expected Output**:
```
Processing 3 students...
✓ Test Student 1: 2 diplomas generated
✓ Test Student 2: 2 diplomas generated
✓ Test Student 3: 2 diplomas generated
Completed: 3/3 students (100%)
```

### Verify Output

```bash
# List generated files
dir output_diplomas\
# Should show:
# Test_Student_1_KZ_2025-2026.xlsx
# Test_Student_1_RU_2025-2026.xlsx
# Test_Student_2_KZ_2025-2026.xlsx
# ... etc
```

---

## Production Configuration

### Directory Structure

```
/diploma_automation/
├── config/           (from git clone)
├── core/             (from git clone)
├── batch/            (from git clone)
├── data/             (from git clone, Phase 2)
├── analysis/         (from git clone)
├── tests/            (from git clone)
├── docs/             (from git clone)
│
├── venv/             (created locally, .gitignored)
│
├── source_files/     (CREATE: source Excel files)
│   ├── 2025-2026_IT_KZ.xlsx
│   ├── 2025-2026_IT_RU.xlsx
│   └── 2025-2026_ACCOUNTING.xlsx
│
├── output_diplomas/  (CREATE: generated files go here)
│   ├── YYYY-MM-DD_batch_1/  (organized by batch date)
│   ├── YYYY-MM-DD_batch_2/
│   └── YYYY-MM-DD_batch_3/
│
├── logs/             (CREATE: log files)
│   ├── diploma_automation.log
│   └── errors.log
│
└── config_prod.ini   (CREATE: environment-specific settings)
```

### Configuration File (Optional)

Create `config_prod.ini` for environment-specific settings:

```ini
[PATHS]
source_file = \\server\shared\2025-2026_grades.xlsx
output_dir = C:\Program Data\Diplomas\output_2025-2026
log_file = C:\Program Data\Diplomas\logs\automation.log
temp_dir = C:\temp\diploma_automation

[PROCESSING]
batch_size = 50
parallel_workers = 4
max_retries = 3
timeout_seconds = 300

[NOTIFICATIONS]
email_on_success = it@university.kz
email_on_failure = it-support@university.kz
slack_webhook = https://hooks.slack.com/services/...

[AUDIT]
log_level = INFO
archive_completed = true
archive_path = \\server\archive\diplomas
keep_days = 30
```

**Load in batch script**:
```python
import configparser

config = configparser.ConfigParser()
config.read('config_prod.ini')

SOURCE_FILE = config['PATHS']['source_file']
OUTPUT_DIR = config['PATHS']['output_dir']
```

---

## Scheduled Automation

### Windows Task Scheduler

**Create scheduled task**:

```powershell
# PowerShell (Run as Administrator)

$action = New-ScheduledTaskAction `
    -Execute "cmd.exe" `
    -Argument "/c cd C:\diploma_automation && venv\Scripts\activate && python -m batch._generate_it >> logs\diploma_automation.log 2>&1"

$trigger = New-ScheduledTaskTrigger `
    -At 22:00 `
    -Daily

Register-ScheduledTask `
    -TaskName "Diploma_Generation_IT" `
    -Action $action `
    -Trigger $trigger `
    -RunLevel Highest
```

**Or via GUI**:
1. Open Task Scheduler
2. Create Basic Task → "Diploma Generation IT"
3. Trigger: Daily at 22:00
4. Action: Start program
5. Program: `cmd.exe`
6. Arguments: `/c cd C:\diploma_automation && venv\Scripts\activate && python -m batch._generate_it`
7. Task will run at 22:00 daily

### Linux Cron Job

```bash
# Edit crontab
crontab -e

# Add line (runs daily at 22:00):
0 22 * * * cd /home/user/diploma_automation && source venv/bin/activate && python -m batch._generate_it >> logs/diploma_automation.log 2>&1
```

---

## Backup Strategy

### Pre-Generation Backup

```bash
# Before batch generation, back up source files
Robocopy "\\server\shared\diplomas" "D:\Backup\diplomas_2025-02-19" /E /COPY:DAT /LOG:backup.log

# Or on Linux:
rsync -av "\\server\shared\diplomas" "/backup/diplomas_2025-02-19/"
```

### Post-Generation Backup

```bash
# After successful generation, archive output
mkdir "D:\Archive\2025-2026_batch_1"
Move-Item "output_diplomas\*" "D:\Archive\2025-2026_batch_1\" -Force
```

### Backup Schedule

| What | When | Where | Retention |
|------|------|-------|-----------|
| Source files | Daily | Network share | 90 days |
| Generated diplomas | After each batch | External drive | 1 year |
| Logs | Daily | Local /logs | 30 days |
| Database (future) | Continuously | SQL backup | 7 days |

---

## Monitoring & Logging

### Enable Logging

Edit [config/settings.py](../config/settings.py):

```python
DEBUG_MODE = False          # Set to True for development
VALIDATE_SCHEMA = True      # Verify Excel structure
LOG_LEVEL = "INFO"          # DEBUG, INFO, WARNING, ERROR
LOG_FILE = "logs/diploma_automation.log"
```

### Log File Locations

```
logs/
├── diploma_automation.log    (all messages)
├── errors.log               (errors only)
└── archive/
    ├── 2025-01-01.log
    ├── 2025-01-02.log
    └── ...
```

### Check Logs

```bash
# View recent entries
tail -f logs/diploma_automation.log

# Count errors
grep "ERROR" logs/diploma_automation.log | wc -l

# Find specific student
grep "Иванов" logs/diploma_automation.log

# Export week's errors
select-string "ERROR" logs/diploma_automation.log | Export-Csv errors_week.csv
```

### Log Rotation

```python
# In batch script
import logging
from logging.handlers import RotatingFileHandler

handler = RotatingFileHandler(
    'logs/diploma_automation.log',
    maxBytes=10485760,  # 10MB
    backupCount=10      # Keep 10 old files
)
logger.addHandler(handler)
```

---

## Troubleshooting

### Issue: "No module named 'config'"

**Cause**: Virtual environment not activated

**Solution**:
```bash
# Windows:
venv\Scripts\activate

# Linux/macOS:
source venv/bin/activate

# Verify (should show "(venv)" in prompt):
(venv) C:\diploma_automation>
```

### Issue: "Permission denied" on output directory

**Cause**: Output directory is read-only or used by another process

**Solution**:
```bash
# Windows: Check file locks
tasklist /M excel.exe  # Is Excel using files?

# Close Excel, then:
icacls output_diplomas /grant %USERNAME%:M

# Linux:
chmod 755 output_diplomas
```

### Issue: "Cannot read source file: Excel is open"

**Cause**: Source file is being edited or viewed

**Solution**:
1. Close the Excel file in all applications
2. Wait 5 seconds
3. Retry generation

Or configure read-only mode:
```python
df = pd.read_excel(SOURCE_FILE, sheet_name='3F-1', engine='openpyxl')
# Automatically handles locked files
```

### Issue: "Memory error" with large batches

**Cause**: Processing 1000+ students with limited RAM

**Solution**: Process in smaller batches
```bash
# Instead of processing all 1000 students at once:
python -m batch._generate_it --start-row 1 --end-row 100
python -m batch._generate_it --start-row 101 --end-row 200
python -m batch._generate_it --start-row 201 --end-row 300
# ... etc
```

Or increase available memory:
```python
import gc
gc.collect()  # Force garbage collection between batches
```

### Issue: "Grades contain special characters"

**Cause**: Non-numeric grades (text, symbols, etc.)

**Solution**: Validate source file
```bash
cd analysis/validators
python validate_excel_structure.py
# Should report issues with non-numeric grades
```

---

## Performance Tuning

### Configuration Optimization

```python
# config/settings.py

# For large batches (1000+ students):
BATCH_SIZE = 50              # Process 50 students at a time
PARALLEL_WORKERS = 4         # Use 4 CPU cores
CACHE_SUBJECT_NAMES = True   # Cache normalized names
ENABLE_PROGRESS_BAR = True   # Show progress in terminal
```

### Profiling

```bash
# Time the generation
python -m cProfile -s cumtime batch/_generate_it.py

# Output will show slowest functions:
# ncalls  tottime  cumtime
# 1000    0.5     5.2    core.converters.convert_score_to_grade
# 500     0.3     3.1    data.excel_generator.render_cell
```

### Database Indexing (Future)

When migrating to database (v2.0):
```sql
CREATE INDEX idx_student_grade ON student_grades(
    student_id,
    subject_id,
    academic_year
);
```

---

## Security Considerations

### Data Protection

```python
# config/settings.py

# 1. Don't log sensitive data
SENSITIVE_FIELDS = ['full_name', 'diploma_number', 'email']
LOG_SENSITIVE = False        # Mask in logs

# 2. Encrypt grades file
ENCRYPT_OUTPUT = True
ENCRYPTION_KEY = "from-environment-variable"

# 3. Access control
REQUIRE_AUTH = True
ADMIN_USERS = ['admin@university.kz']
```

### File Permissions

```bash
# Windows: Restrict to administrators
icacls "C:\diploma_automation\config" /grant:r SYSTEM:F /grant:r Administrators:F /remove Users

# Linux: Restrict to owner
chmod 700 config/
chmod 600 config/settings.py
```

### Audit Trail

```python
import logging
import json
from datetime import datetime

def audit_log(action, student, status, error=None):
    """Log all diploma generation actions for audit."""
    log_entry = {
        "timestamp": datetime.now().isoformat(),
        "action": action,
        "student_id": student.diploma_number,
        "status": status,
        "error": error
    }
    with open("logs/audit.log", "a") as f:
        f.write(json.dumps(log_entry) + "\n")

# Usage:
audit_log("generate", student, "success")
audit_log("generate", student, "failed", "grade_validation_error")
```

---

## Update & Maintenance

### Updating Code

```bash
# Pull latest changes
git pull origin main

# Reinstall in case dependencies changed
pip install -r requirements.txt --upgrade

# Run tests to verify
pytest tests/ -v

# If all pass, safe to resume operations
```

### Version Control

```bash
# Keep production config out of version control:
echo "config_prod.ini" >> .gitignore
echo "logs/" >> .gitignore
echo "output_diplomas/" >> .gitignore
echo "*.xlsx" >> .gitignore

git add .gitignore
git commit -m "Add production files to .gitignore"
```

### Database Migration (Future)

When upgrading from v1.x to v2.0 (with database):
```bash
# 1. Backup old files
python scripts/backup_old_system.py

# 2. Migrate data
python scripts/migrate_to_database.py

# 3. Verify migration
python scripts/verify_migration.py

# 4. Run compatibility tests
pytest tests/test_migration.py -v
```

---

## Disaster Recovery

### Complete System Restore

```bash
# 1. Restore from backup
robocopy "D:\Backup\system_2025-02-19" "C:\diploma_automation" /E

# 2. Reinstall Python environment
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt

# 3. Restore configuration
copy "D:\Backup\config_prod.ini" "C:\diploma_automation\config_prod.ini"

# 4. Verify system
python -c "from config import settings; settings.validate()"

# 5. Test with sample
python -c "from core.converters import convert_score_to_grade; convert_score_to_grade('85')"
```

### Quick Recovery Procedures

| Scenario | Recovery Time | Steps |
|----------|---|---|
| Single file corrupted | 5 min | Restore from backup, retry |
| Python environment broken | 15 min | Reinstall venv, reinstall deps |
| Config file lost | 10 min | Restore from backup, update paths |
| Database (future) lost | 30 min | Restore from SQL backup, verify integrity |
| Entire system loss | 1 hour | Restore from full backup, reconfigure paths, test |

---

## Success Criteria

### Installation Complete When:

- ✅ `python -c "from config import settings"` succeeds
- ✅ `python -c "from core.converters import convert_score_to_grade"` succeeds
- ✅ Test file generated with 3+ students
- ✅ 6+ output diplomas created (3 students × 2 languages)
- ✅ Output files are valid Excel (.xlsx format)
- ✅ Diplomas are readable and correctly formatted

### Monitoring Active When:

- ✅ Logs being written to logs/diploma_automation.log
- ✅ Daily backup running on schedule
- ✅ Email alerts configured for errors
- ✅ Memory usage < 1GB for typical batch
- ✅ Generation time < 2 seconds per diploma

### Production Ready When:

- ✅ All tests passing (`pytest tests/ -v`)
- ✅ 100+ student batch processed successfully
- ✅ Two consecutive successful generations on schedule
- ✅ Backup & restore tested successfully
- ✅ Team trained on troubleshooting procedures

---

## Support Contacts

| Issue | Contact | Response Time |
|-------|---------|---|
| Installation help | it-support@university.kz | 4 hours |
| System down | ops-team@university.kz | 1 hour |
| Code bug | dev-team@university.kz | 8 hours |
| Data quality | registrar@university.kz | 24 hours |
| Security incident | security@university.kz | 30 min |

---

## Change Log

| Version | Date | Changes |
|---------|------|---------|
| 1.0 | 2025-02-19 | Initial deployment guide |
| 1.1 | TBD | Add database backup procedures |
| 2.0 | TBD | Web interface deployment |

---

*Last updated: February 19, 2025*
*Next review: March 19, 2025 (monthly)*
