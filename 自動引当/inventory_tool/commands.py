from __future__ import annotations

import click
from flask import Flask

from .sample_data import seed_demo_data
from .services import recalculate_allocations, run_checks


def register_commands(app: Flask) -> None:
    @app.cli.command("seed-demo")
    def seed_demo_command() -> None:
        """Create demo data for the inventory tool."""
        seed_demo_data()
        click.echo("デモデータを投入しました。")

    @app.cli.command("recalculate")
    def recalculate_command() -> None:
        """Recalculate allocation and order status."""
        recalculate_allocations()
        click.echo("引当再計算を完了しました。")

    @app.cli.command("run-checks")
    def run_checks_command() -> None:
        """Run alert checks."""
        run_checks()
        click.echo("漏れチェックを完了しました。")
