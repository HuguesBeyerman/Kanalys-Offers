import click

@click.command()
@click.option(
    "--weekday",
    prompt='walue',
    type=click.INT
)
def cli(weekday):
    click.echo(f"Weekday: {weekday}")

if __name__ == "__main__":
    cli()