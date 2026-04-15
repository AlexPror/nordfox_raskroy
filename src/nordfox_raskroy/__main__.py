import sys


def main() -> None:
    if len(sys.argv) > 1 and sys.argv[1] in ("--version", "-V"):
        from nordfox_raskroy import __version__

        print(__version__)
        return
    from nordfox_raskroy.app import run_app

    run_app()


if __name__ == "__main__":
    main()
