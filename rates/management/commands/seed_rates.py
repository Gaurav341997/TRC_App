from django.contrib.auth.models import User
from django.core.management.base import BaseCommand
from django.db import transaction

from rates.models import Product, Sheet, Group, Slab, Tenure, Rate

SLAB_DEFS = [
    # label, amt_from, amt_to, group_index, tint
    ("3cr to <3.25cr", 30000000, 32499999, 0, False),
    ("3.25cr to <3.50cr", 32500000, 34999999, 0, True),
    ("3.50cr to 10cr", 35000000, 100000000, 0, False),
    ("Above 10cr to 15cr", 100000001, 150000000, 1, False),
    ("Above 15 cr to 20 cr", 150000001, 200000000, 1, True),
    ("Above 20 cr to 25 cr", 200000001, 250000000, 1, False),
    ("Above 25 cr to 35 cr", 250000001, 350000000, 1, True),
    ("Above 35 cr to 50 cr", 350000001, 500000000, 1, False),
    ("Above 50 cr to 70 cr", 500000001, 700000000, 2, True),
    ("Above 70 cr to 95 cr", 700000001, 950000000, 2, False),
    ("Above 95 cr to 105 cr", 950000001, 1050000000, 2, True),
    ("Above 105 cr to 150 cr", 1050000001, 1500000000, 3, False),
    ("Above 150 cr to 200 cr", 1500000001, 2000000000, 3, True),
    ("Above 200 cr", 2000000001, 99999999999, 3, False),
]
TENURE_DEFS = [(0, 7), (1, 1), (1, 16), (2, 1), (2, 29), (3, 1), (4, 1), (6, 1), (9, 1), (12, 1), (12, 5), (12, 6)]
CALLABLE_GRID = [
    ["5.00", "5.00", "5.00", "5.00", "5.00", "5.00", "5.00", "5.00", "5.00", "5.00", "5.00", "5.00", "5.00", "5.00"],
    ["5.50", "5.50", "5.50", "5.50", "5.50", "5.50", "5.50", "5.50", "5.50", "5.50", "5.50", "5.50", "5.50", "5.50"],
    ["5.80", "5.75", "5.80", "5.80", "5.75", "5.80", "5.75", "5.80", "5.80", "5.75", "5.80", "5.75", "5.80", "5.80"],
    ["6.25", "5.75", "6.25", "6.25", "5.75", "6.25", "5.75", "6.25", "6.25", "5.75", "6.25", "5.75", "6.25", "6.25"],
    ["5.65", "5.65", "5.65", "5.65", "5.65", "5.65", "5.65", "5.65", "5.65", "5.65", "5.65", "5.65", "5.65", "5.65"],
    ["6.65", "5.85", "6.65", "6.65", "5.85", "6.65", "5.85", "6.65", "6.65", "5.85", "6.65", "5.85", "6.65", "6.65"],
    ["6.90", "5.85", "6.90", "6.90", "5.85", "6.90", "5.85", "6.90", "6.90", "5.85", "6.90", "5.85", "6.90", "6.90"],
    ["7.00", "5.90", "7.00", "7.00", "5.90", "7.00", "5.90", "7.00", "7.00", "5.90", "7.00", "5.90", "7.00", "7.00"],
    ["7.20", "5.90", "7.20", "7.20", "5.90", "7.20", "5.90", "7.20", "7.20", "5.90", "7.20", "5.90", "7.20", "7.20"],
    ["7.40", "5.90", "7.40", "7.40", "5.90", "7.40", "5.90", "7.40", "7.40", "5.90", "7.40", "5.90", "7.40", "7.40"],
    ["7.40", "5.90", "7.40", "7.50", "5.90", "7.50", "5.90", "7.50", "7.50", "5.90", "7.50", "5.90", "7.50", "7.50"],
    ["7.40", "5.90", "7.40", "7.60", "5.90", "7.60", "5.90", "7.60", "7.60", "5.90", "7.60", "5.90", "7.60", "7.60"],
]


def grid_for(slug):
    if slug in ("noncallable", "ib_noncallable"):
        return [[f"{float(v) - 0.25:.2f}" for v in row] for row in CALLABLE_GRID]
    elif slug == "fin_callable":
        return [[f"{float(v) + 0.15:.2f}" for v in row] for row in CALLABLE_GRID]
    return CALLABLE_GRID


def build_sheet(product, kind, sub):
    sheet, _ = Sheet.objects.get_or_create(product=product, kind=kind, defaults={"effective_date": "01-Sep-26"})
    sheet.effective_date = "01-Sep-26"
    sheet.save()
    sheet.rates.all().delete()
    sheet.slabs.all().delete()
    sheet.tenures.all().delete()
    sheet.groups.all().delete()

    groups = [
        Group.objects.create(sheet=sheet, label="3 -10 crs", order=0),
        Group.objects.create(sheet=sheet, label="10-50crs", order=1),
        Group.objects.create(sheet=sheet, label="50-105crs", order=2),
        Group.objects.create(sheet=sheet, label="Greater than 105", order=3),
    ]
    slabs = []
    for i, (label, af, at, gi, tint) in enumerate(SLAB_DEFS):
        slabs.append(Slab.objects.create(sheet=sheet, group=groups[gi], label=label,
                                         amt_from=af, amt_to=at, tint=tint, sub=sub, order=i))
    tenures = [Tenure.objects.create(sheet=sheet, from_month=m, to_days=d, order=i)
               for i, (m, d) in enumerate(TENURE_DEFS)]

    grid = grid_for(product.slug)
    rates = []
    for r, t in enumerate(tenures):
        for c, s in enumerate(slabs):
            rates.append(Rate(sheet=sheet, tenure=t, slab=s, value=grid[r][c]))
    Rate.objects.bulk_create(rates)


class Command(BaseCommand):
    help = "Seed users, products and rate sheets."

    @transaction.atomic
    def handle(self, *args, **opts):
        from django.contrib.auth.models import Permission

        if not User.objects.filter(username="admin").exists():
            User.objects.create_user("admin", password="admin123", is_staff=True, is_superuser=True)
        else:
            admin_user = User.objects.get(username="admin")
            admin_user.is_staff = True
            admin_user.is_superuser = True
            admin_user.set_password("admin123")
            admin_user.save()

        if not User.objects.filter(username="user").exists():
            demo_user = User.objects.create_user("user", password="user123", is_staff=False)
        else:
            demo_user = User.objects.get(username="user")
            demo_user.set_password("user123")
            demo_user.save()

        for codename in ("access_trc", "save_trc_draft"):
            perm = Permission.objects.filter(codename=codename).first()
            if perm:
                demo_user.user_permissions.add(perm)

        for slug, label, sub in [
            ("callable", "Non Fin Callable", "Callable"),
            ("noncallable", "Non Fin / Fin Non-Callable", "Non-Callable"),
            ("ib_callable", "IB Callable", "Callable"),
            ("ib_noncallable", "IB Non-Callable", "Non-Callable"),
            ("fin_callable", "Financial Callable", "Callable"),
        ]:
            product, _ = Product.objects.get_or_create(slug=slug, defaults={"label": label})
            product.label = label
            product.save()
            build_sheet(product, Sheet.PUBLISHED, sub)
            build_sheet(product, Sheet.WORKING, sub)

        self.stdout.write(self.style.SUCCESS("Seeded users (admin/admin123, user/user123) and both products."))
