# Bilant Generator

A Python web application that generates Romanian Balance Sheets (Bilanț) from Trial Balances (Balanță de Verificare).

## Features

- Upload Excel file with Balanta and Bilant sheets
- Automatic calculation of account balances based on formulas
- Support for complex formulas: `+`, `-`, `+/-` (dynamic sign), `dinct.` (from account)
- Verification breakdown showing all account contributions
- Download generated Excel file with results

## Local Development

### Prerequisites

- Python 3.9+
- pip

### Installation

```bash
# Clone the repository
git clone https://github.com/YOUR_USERNAME/Bilant-Generator.git
cd Bilant-Generator

# Create virtual environment
python3 -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt

# Run the application
python app.py
```

The app will be available at `http://localhost:5000`

## Deployment to DigitalOcean App Platform

### Option 1: Using GitHub Integration

1. **Push to GitHub**
   ```bash
   cd Bilant-Generator
   git init
   git add .
   git commit -m "Initial commit"
   gh repo create Bilant-Generator --public --source=. --push
   ```

2. **Create DigitalOcean App**
   - Go to [DigitalOcean App Platform](https://cloud.digitalocean.com/apps)
   - Click "Create App"
   - Select "GitHub" as source
   - Choose your repository
   - Select the branch (main)
   - DigitalOcean will auto-detect the Dockerfile
   - Choose instance size (Basic $5/month is sufficient)
   - Click "Create Resources"

3. **Update app.yaml** (optional)
   - Edit `.do/app.yaml` with your GitHub username
   - This allows one-click deployment

### Option 2: Using DigitalOcean CLI

```bash
# Install doctl
brew install doctl  # macOS

# Authenticate
doctl auth init

# Create app from spec
doctl apps create --spec .do/app.yaml
```

### Environment Variables (if needed)

No environment variables are required for basic operation.

## File Structure

```
Bilant-Generator/
├── app.py              # Main Flask application
├── templates/
│   └── index.html      # Upload interface
├── requirements.txt    # Python dependencies
├── Dockerfile          # Container configuration
├── .do/
│   └── app.yaml        # DigitalOcean App Platform spec
├── .gitignore
└── README.md
```

## Input File Format

The uploaded Excel file must contain:

### Sheet: "Balanta"
| Column | Content |
|--------|---------|
| B | Account number (RAD1) |
| E | SFD (Debit balance) |
| F | SFC (Credit balance) |

### Sheet: "Bilant"
| Column | Content |
|--------|---------|
| A | Description with formula (e.g., "Terenuri (ct. 211+212-2811)") |
| C | Row number (Nr. rd.) |
| E | Calculated value (output) |

## Formula Syntax

The application supports the following formula patterns:

- **Simple addition**: `201+202+203`
- **Subtraction**: `201-2801-2901`
- **Dynamic sign** (`+/-`): `345+346+/-348` - For accounts that can be debit or credit
- **From account** (`dinct.`): `-dinct.4428` - Subtract specific account balance
- **Row references**: `TOTAL (rd. 01 la 06)` - Sum of rows 1 through 6

## Output

The generated Excel file contains:

1. **Balanta sheet**: Original data + calculated Sold Final
2. **Bilant sheet**: Original data + calculated values + verification details

## Technical Notes

- The application fixes the VBA macro bug with `+/-` parsing (off-by-one error)
- Decimal values are preserved (no truncation)
- All accounts matching a prefix are included in verification

## License

MIT License
