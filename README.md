# invoice

Skript pro generování měsíčních faktur AnChoice ze šablony v Excelu.

## Požadavky

- Windows s nainstalovaným Microsoft Excel (skript ovládá Excel přes COM, aby zachoval
  přesné formátování šablony).
- [uv](https://docs.astral.sh/uv/)

## Použití

```
uv run create_invoice.py <mesic 1-12> [--rok ROK] [--prepsat]
```

Příklad — faktura za srpen, splatná v září:

```
uv run create_invoice.py 9
```

Vytvoří se `data/anc_{rok}{mesic:03d}.xls`, např. `data/anc_2026009.xls`, a zkopíruje se do
`H:\My Drive\Work\Petr Erben\Faktury\{rok}`. Pokud tam faktura se stejným názvem už existuje,
skript ji nepřepíše (pro přepsání použij `--prepsat`).

Automaticky dopočítaná pole:

- **číslo faktury / variabilní symbol**: `{rok}{mesic:03d}`
- **datum vystavení**: dnešní datum
- **DUZP**: poslední den měsíce předcházejícího zadanému měsíci
- **splatnost**: 15. den zadaného měsíce
- **položka "Programátorské práce..."**: za měsíc a rok odpovídající DUZP

Šablona je v `templates/anc_sablona.xls`. Vygenerované faktury v `data/` nejsou verzované
v gitu (obsahují reálné fakturační údaje) — pokud je chceš mít v repozitáři, uprav `.gitignore`.

### Odeslání faktury emailem

```
uv run send_invoice.py <mesic 1-12> [--rok ROK]
```

Faktura musí už existovat v `data/` (spusť nejdřív `create_invoice.py`). Skript ji převede
na PDF (`data/anc_{cislo}.pdf`) a pošle emailem z `petr@anchoice.cz` na `erbenpet@gmail.com`.

Potřebuje SMTP údaje v `.env` v kořeni repozitáře (zkopíruj `.env.example` a vyplň):

```
SMTP_HOST=...
SMTP_PORT=587
SMTP_USER=petr@anchoice.cz
SMTP_PASS=...
```

`.env` je v `.gitignore`, takže se do repozitáře nedostane.
