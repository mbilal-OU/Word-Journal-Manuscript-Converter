from word_journal_manuscript_converter.cli import build_parser


def test_cli_parses_inspect():
    args = build_parser().parse_args(["inspect", "paper.docx"])
    assert args.command == "inspect"
    assert args.docx == "paper.docx"
