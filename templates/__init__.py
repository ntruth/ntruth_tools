"""Helpers for accessing the built-in Excel template."""
from __future__ import annotations

import base64
from pathlib import Path
from typing import Optional

DEFAULT_TEMPLATE_FILENAME = "txt_to_excel_template.xlsx"

# Base64 representation of the built-in Excel template.  The template is a
# minimal workbook with a single worksheet and no extra styling.  It exists so
# that users always have a starting point that matches the behaviour described
# in the original requirements document.
DEFAULT_TEMPLATE_B64 = (
    "UEsDBBQAAAAIAGJtXlt67VFpGQEAADgDAAATAAAAW0NvbnRlbnRfVHlwZXNdLnhtbK1SS08CMRC+"
    "8yuaXgkteDDG7MLBx1E94A8Yu7NsQ1/pDAj/3lLUGKOGRHpo03zPTKZZ7LwTW8xkY2jlTE2lwGBi"
    "Z8Oqlc/L+8mVFMQQOnAxYCv3SHIxHzXLfUISRRyolQNzutaazIAeSMWEoSB9zB64fPNKJzBrWKG+"
    "mE4vtYmBMfCEDx5yPhLlNLfYw8axuNsV7NgmoyMpbo7sQ2ArISVnDXDB9TZ036Im7zGqKCuHBpto"
    "XAhS/x5zgH9P+Sp+LIPKtkPxBJkfwBeq3jn9GvP6Jca1+tvph76x763BLpqNLxJFKSN0NCCyd6q+"
    "yoMN4xNLVAXp+szO3ObT/5QyxHuHdO55VNNT4rnsIR7v/4+h2nykNrou/vwNUEsDBBQAAAAIAGJt"
    "Xlte6cK7qgAAAB0BAAALAAAAX3JlbHMvLnJlbHONj8EOgjAQRO98RbN3KXgwxli4GBOuBj+glqUQ"
    "aLdpq+Lf26MYD85tM7NvMsd6MTN7oA8jWQFlXgBDq6gbrRZwbc+bPdRVdrzgLGOKhGF0gaUfGwQM"
    "MboD50ENaGTIyaFNTk/eyJhOr7mTapIa+bYodtx/MqDKWNIKzJpOgG+6Elj7cvhPAfX9qPBE6m7Q"
    "xh89X4lEll5jFLDM/El+uhFNeYICTyv5amb1BlBLAwQUAAAACABibV5bQHqjy7YAAAAfAQAADwAA"
    "AHhsL3dvcmtib29rLnhtbI2PTQ6CQAyF955i0r0MuDCGAG6MiWv1ACMUmcBMSTv+HN8RNG7t6r3X"
    "5mtbbJ9uUHdkseRLyJIUFPqaGuuvJZxP++UGttWieBD3F6JexXEvJXQhjLnWUnfojCQ0oo+dltiZ"
    "EC1ftYyMppEOMbhBr9J0rZ2xHmZCzv8wqG1tjTuqbw59mCGMgwnxWOnsKFAtVKxiWiOz+QXKG4cl"
    "HN86AzVlhyZ+CYpzGwUfmgz0h6G/kEJ/v61eUEsDBBQAAAAIAGJtXls76viN0AAAACsCAAAaAAAA"
    "eGwvX3JlbHMvd29ya2Jvb2sueG1sLnJlbHOtkcuKAjEQRfd+Rai9Xd0ODMPQaTciuBXnA0K6+oHd"
    "SUiVOv79BAUfIKMLsyluVereQ1LOf8dB7Sly752GIstBkbO+7l2r4WeznH7BvJqUaxqMpCvc9YFV"
    "2nGsoRMJ34hsOxoNZz6QS5PGx9FIkrHFYOzWtISzPP/EeOsB1USlc2esVrWGuKoLUJtjoFcCfNP0"
    "lhbe7kZy8iAHDz5uuSOSZGpiS6Lh0mI8lSJLroD/EM3eScRyHIivOGf9lOHjnQySdumKcJLn5uUx"
    "Srz79OoPUEsDBBQAAAAIAGJtXluTkavGmQAAAOAAAAAYAAAAeGwvd29ya3NoZWV0cy9zaGVldDEu"
    "eG1sjY5JDsIwDEX3PUWUPU1hgVDVYVNxAjiAlTpNRBNHcRiOT9RKrPHuy/b7rxs/fhUvTOwo9PJY"
    "N1Jg0DS7sPTyfrseLnIcqu5N6cEWMYtyH7iXNufYKsXaogeuKWIoG0PJQy4xLYpjQpi3J7+qU9Oc"
    "lQcX5E5o0z8MMsZpnEg/PYa8QxKukIstWxdZDpUo0201E2RQRVb9bIcvUEsDBBQAAAAIAGJtXlud"
    "JnSIYwEAAPcDAAANAAAAeGwvc3R5bGVzLnhtbK1TPW+DMBDd8yss740TpFZVBclQKVKXLkmlrgYM"
    "WDrbyL5Eob++JoQADhkq1dPdvbv3nr/i7VkBOQnrpNEJXS9XlAidmVzqMqFfh93TK91uFrHDBsS+"
    "EgKJH9AuoRVi/caYyyqhuFuaWmiPFMYqjj61JXO1FTx37ZACFq1WL0xxqelmQfyKC6PRkcwcNXrh"
    "a/WGDOml5H7IiYPvW1MWQJkBYwl6G6LlCWHNlehm3znI1Mq7joIrCU3XE43RmA1Outj13iXAzXs0"
    "8e6RgL/miMLqnUfINT40tTerjRaB3mT6T2Sl5c06en7E18W9/9TY3N/57Ol3WCALosDw3Kwsq7si"
    "mjospQbRqLCaS14azWHid6zdZ73nTADs22f4XcwbPxdEH9VO4UeeUP+O2wvrQ7/3a9iRdgnrlcbc"
    "I7n/UyLnYkZyqnYxMC94g0n7nBP62X4zGFhJepSAUj/alpeJ2fCHN79QSwMEFAAAAAgAYm1eW+yx"
    "5Az/AwAAyh8AABMAAAB4bC90aGVtZS90aGVtZTEueG1s7VlLb9s4EL73VxC8t7KsR+ygSlG7Fvaw"
    "iy2aLPZMS9RjS1ECySTNv19KsmzZJinZaN2kjQ6GTQ6/mfnIGc7I7z98Kwh4wIznJQ2g/W4CAaZR"
    "Gec0DeA/d+HbGQRcIBojUlIcwCfM4YebN+/RtchwgYFcTvk1CmAmRHVtWTySw4i/KytM5VxSsgIJ"
    "+ZOlVszQo4QtiDWdTHyrQDmFgKJCov6dJHmEwV0NCW/eAPl0GlZEflDB29HNTETYbdTo76+HO5mN"
    "XPzV3h/bjPMnviQMPCASQGlTXD7e4W8CAoK4kBMBnDQPtA4QLQWkhCNirJqeirB5VCqO4BpPpmoV"
    "LF1vddihO7/6pLZ6qrB6BORqtVqubLWVx5AoiuReacjow7rhzF4oLVVCbKFHWLyceBPXAK2z2hmG"
    "ni8WC2+uh3Y00O4w9Gziux+nemhXA+2N4Hrxcbn09dCeBtofhg6v5r5rgPaPoDOS06/DwHX0qYND"
    "ASAXJyX5YxzyTCLPlJF9jFGPblPNXgJKSipGZaAC/VeyUEorLSNI5BSIpwonKJIoS0TyNcsPrdtI"
    "Y9QT1chE3CBT+6MxqDY1p8/M1GOD2n3quN/fkmLcjiQ5IbfiieA/udpTXpI8DqXU8Wwn0ejZHqkq"
    "k1+VTlpGNImUMtRMAlaKf3OR3WaokrbbUKs75Uq79yRAVXIZQxqQIUfMqzYr5WEQ7SJPdVPq9CHx"
    "Vxm36xzlFXu0yOpZaXDbqv0eRYtTG3wZapyrl0WN3er9gdyQ++JELzuVPXrsS9HTzqnzhLUL3l81"
    "ssfx3Cnsn99TckL/SNjTUzRmKMadpd6ZR2mUwu8fab57sSR0bn62nYtvovOyNvHHp8vvcMtePhSf"
    "+y6eldWbwlNbOdbcUWNRSSh4lA2kN/UgiFAVwEQW1PJrUcUB5DSFAJGUBjASuqMxXJl2UiOr086t"
    "4Xq3Ylx8QjxrARt5LWBdtgvMAMmLttlS72nzBoEaqJp6MkW+cjWKK2dmv3Kl56od00YuThIcCWP0"
    "9kS05rUyQyVceS+duM3iR7Am9+wLkrvkNr6AOOdCHvvuB5M53zUWYLVbBqW7aY3dl/XLefXL7NdG"
    "IY8wxU5sNCpCBWYI1OETwJKJrJSXVpXlUchKKkxXaJ0y8jQTX/IUsDwNoMgYxp/FxgMxXMzIdgbI"
    "FBPAt+2Fbbyxm8jb6DMkCpPLNSXVAB9r/IDJXZ0J/bregSDr7g9TelKiGk/XwaQqmazT8Dm9XzpT"
    "l5HtXVE6H+q6Dku8q6H4MJR2v8XLtPm5xfNJ/cGuWp/PzuwOJtOf0nJ5l+u4TqHm7L65vxXnNmqj"
    "Gq6X2P32yPF/2aZSd2E0DWex90dHPXTwn/h26OZ/UEsBAhQDFAAAAAgAYm1eW3rtUWkZAQAAOAMA"
    "ABMAAAAAAAAAAAAAAIABAAAAAFtDb250ZW50X1R5cGVzXS54bWxQSwECFAMUAAAACABibV5bXunC"
    "u6oAAAAdAQAACwAAAAAAAAAAAAAAgAFKAQAAX3JlbHMvLnJlbHNQSwECFAMUAAAACABibV5bQHqj"
    "y7YAAAAfAQAADwAAAAAAAAAAAAAAgAEdAgAAeGwvd29ya2Jvb2sueG1sUEsBAhQDFAAAAAgAYm1e"
    "Wzvq+I3QAAAAKwIAABoAAAAAAAAAAAAAAIABAAMAAHhsL19yZWxzL3dvcmtib29rLnhtbC5yZWxz"
    "UEsBAhQDFAAAAAgAYm1eW5ORq8aZAAAA4AAAABgAAAAAAAAAAAAAAIABCAQAAHhsL3dvcmtzaGVl"
    "dHMvc2hlZXQxLnhtbFBLAQIUAxQAAAAIAGJtXludJnSIYwEAAPcDAAANAAAAAAAAAAAAAACAAdcE"
    "AAB4bC9zdHlsZXMueG1sUEsBAhQDFAAAAAgAYm1eW+yx5Az/AwAAyh8AABMAAAAAAAAAAAAAAIAB"
    "ZQYAAHhsL3RoZW1lL3RoZW1lMS54bWxQSwUGAAAAAAcABwDBAQAAlQoAAAAA"
)


def default_template_bytes() -> bytes:
    """Return the decoded bytes for the built-in template."""

    return base64.b64decode(DEFAULT_TEMPLATE_B64)


def ensure_default_template_file(path: Optional[Path] = None) -> Path:
    """Ensure the default template exists on disk and return its path.

    The template is written lazily the first time this function is called so
    that users can still reference it by path (e.g. in the GUI file chooser)
    without shipping a binary file in the repository.
    """

    if path is None:
        path = Path(__file__).with_name(DEFAULT_TEMPLATE_FILENAME)

    if not path.exists():
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(default_template_bytes())

    return path
