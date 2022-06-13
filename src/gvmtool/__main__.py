import click
import gvmtool.report


def print_help(ctx, param, value):
    if value is True:
        click.echo(ctx.get_help())
        ctx.exit()


@click.group()
def main():
    pass


@main.command()
@click.option('--path', default=None, help='csv file path')
@click.option('--out', default=None, help='output filename except the extension')
@click.pass_context
def merge_report(ctx, path, out):
    print_help(ctx, None, value=path is None)
    gvmtool.report.merge_report(path, out)


if __name__ == '__main__':
    main()
