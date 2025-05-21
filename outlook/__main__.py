import typer
from outlook.win import get_emails as win_get_emails, get_email as win_get_email
from rich.table import Table
from rich.console import Console
from rich.panel import Panel
from rich.markdown import Markdown


app = typer.Typer()
console = Console()


def shorten(text, width=30):
    if not text:
        return ""
    return text if len(text) <= width else text[: width - 3] + "..."


@app.command()
def get_emails(
    hours: int = typer.Option(7, help="Number of hours to look back (default: 7)")
):
    """Get emails from Outlook"""

    email_list = win_get_emails(max_hours=hours)
    if not email_list:
        typer.echo("No emails found.")
        return

    table = Table(show_header=True, header_style="bold magenta")
    table.add_column("No", style="dim", width=4)
    table.add_column("Received At", style="dim", width=17)
    table.add_column("Subject", width=40)
    table.add_column("Sender", width=20)
    table.add_column("Sender Email", width=25)
    table.add_column("To", width=20)

    for idx, email in enumerate(email_list, 1):
        table.add_row(
            str(idx),
            email.received_at.strftime("%Y-%m-%d %H:%M") if email.received_at else "",
            shorten(email.subject, 40),
            shorten(email.sender_name, 20),
            shorten(email.sender_email, 25),
            shorten(email.to, 20),
        )
    console.print(table)

    # Prompt user to select an email
    try:
        choice = typer.prompt(
            "Enter the number of the email to view details (or press Enter to exit)",
            default="",
            show_default=False,
        )
        if not choice.strip():
            return

        idx = int(choice)
        if idx < 1 or idx > len(email_list):
            typer.echo("Invalid selection.")
            return

        email = win_get_email(email_list[idx - 1].identifier)
        if not email:
            typer.echo("Email not found.")
            return

        # Header Table
        detail_table = Table(show_header=False, box=None)
        detail_table.add_row("[bold]Subject[/bold]", email.subject or "")
        detail_table.add_row(
            "[bold]From[/bold]", f"{email.sender_name} <{email.sender_email}>"
        )
        detail_table.add_row(
            "[bold]Received At[/bold]",
            str(email.received_at) if email.received_at else "",
        )
        detail_table.add_row("[bold]To[/bold]", email.to or "")
        if email.cc:
            detail_table.add_row("[bold]CC[/bold]", email.cc)
        if email.attachments:
            detail_table.add_row(
                "[bold]Attachments[/bold]",
                ", ".join([a.filename for a in email.attachments]),
            )
        console.print(Panel(detail_table, title="Email Info", expand=False))
        # Body
        if email.body:
            console.print(Panel(email.body, title="Body", expand=True))
        else:
            console.print(
                Panel("[italic]No body available.[/italic]", title="Body", expand=True)
            )
    except (ValueError, KeyboardInterrupt):
        typer.echo("No email selected.")


@app.command()
def send_email(
    to: str = typer.Option(..., help="Recipient email address"),
    subject: str = typer.Option(..., help="Email subject"),
    body: str = typer.Option(..., help="Email body"),
):
    """Send an email through Outlook"""
    typer.echo(f"Sending email to: {to}")
    typer.echo(f"Subject: {subject}")
    typer.echo(f"Body: {body}")
    # TODO: Implement actual email sending logic here


if __name__ == "__main__":
    app()
