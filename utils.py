def has_hyperlink(run):
    try:
        return bool(run.hyperlink and run.hyperlink.address)
    except KeyError:
        return True
