"""
generate_deck.py — CLI entry point for the M3 deck automation.
"""
import argparse
import os
import re
import sys

from m3_deck_automation.excel_reader import read_excel
from m3_deck_automation.deck_writer import generate_deck, TEMPLATE_PATH
from m3_deck_automation.reference_data import load_reference_data

PROJECT_DIR = os.path.dirname(os.path.abspath(__file__))


def safe_filename(name):
    return re.sub(r'[^\w\s-]', '', name).strip().replace(' ', '_')


def parse_args():
    parser = argparse.ArgumentParser(description="Generate M3 portfolio transition deck")
    parser.add_argument("--client", help="Client full name (prompted if omitted)")
    parser.add_argument("--excel", required=True, help="Path to client Excel workbook")
    parser.add_argument("--output", help="Output .pptx path (auto-derived if omitted)")
    return parser.parse_args()


def main():
    args = parse_args()
    client_name = args.client
    if not client_name:
        client_name = input("Client name: ").strip()
    if not client_name:
        print("ERROR: client name required", file=sys.stderr)
        sys.exit(1)

    output_path = args.output or f"{safe_filename(client_name)}_deck.pptx"

    print(f"Client:   {client_name}")
    print(f"Excel:    {args.excel}")
    print(f"Template: {TEMPLATE_PATH}")
    print(f"Output:   {output_path}")
    print()

    print("Reading Excel...")
    excel_data = read_excel(args.excel)
    print()

    print("Loading reference data...")
    ref_data, rr_category = load_reference_data(PROJECT_DIR)
    print()

    print("Generating deck...")
    generate_deck(TEMPLATE_PATH, excel_data, client_name, output_path, ref_data, rr_category)
    print()
    print("Done.")


if __name__ == "__main__":
    main()
