from django.db import models


class Product(models.Model):
    slug = models.SlugField(unique=True)
    label = models.CharField(max_length=100)

    def __str__(self):
        return self.label


class Sheet(models.Model):
    PUBLISHED = "published"
    WORKING = "working"

    product = models.ForeignKey(
        Product,
        on_delete=models.CASCADE,
        related_name="sheets"
    )

    kind = models.CharField(max_length=10)
    effective_date = models.CharField(
        max_length=20,
        default="06-Aug-26"
    )

    class Meta:
        unique_together = ("product", "kind")
        permissions = [
            ("access_trc", "Can access Treasury Rate Console"),
            ("save_trc_draft", "Can save TRC draft rates"),
            ("publish_trc_rates", "Can publish TRC rates"),
            ("upload_trc_rates", "Can upload TRC rates"),
            ("edit_trc_published", "Can edit TRC published rates"),
            ("access_rate_console", "Can access Module: Rate Console"),
            ("access_retail_td", "Can access Module: Retail TD Comparison"),
            ("access_bulk_td", "Can access Module: Bulk TD Comparison"),
            ("access_nre_nro_td", "Can access Module: NRE/NRO TD Comparison"),
        ]

    def __str__(self):
        return f"{self.product.slug}:{self.kind}"


class Group(models.Model):
    sheet = models.ForeignKey(
        Sheet,
        on_delete=models.CASCADE,
        related_name="groups"
    )

    label = models.CharField(max_length=100)
    order = models.IntegerField(default=0)


class Slab(models.Model):
    sheet = models.ForeignKey(
        Sheet,
        on_delete=models.CASCADE,
        related_name="slabs"
    )

    group = models.ForeignKey(
        Group,
        on_delete=models.CASCADE,
        related_name="slabs"
    )

    label = models.CharField(max_length=120)
    amt_from = models.BigIntegerField(
        null=True,
        blank=True
    )

    amt_to = models.BigIntegerField(
        null=True,
        blank=True
    )

    tint = models.BooleanField(default=False)

    sub = models.CharField(
        max_length=40,
        default="Callable"
    )

    order = models.IntegerField(default=0)


class Tenure(models.Model):
    sheet = models.ForeignKey(
        Sheet,
        on_delete=models.CASCADE,
        related_name="tenures"
    )

    from_month = models.IntegerField(default=0)
    to_days = models.IntegerField(default=0)
    order = models.IntegerField(default=0)


class Rate(models.Model):
    sheet = models.ForeignKey(
        Sheet,
        on_delete=models.CASCADE,
        related_name="rates"
    )

    tenure = models.ForeignKey(
        Tenure,
        on_delete=models.CASCADE
    )

    slab = models.ForeignKey(
        Slab,
        on_delete=models.CASCADE
    )

    value = models.CharField(
        max_length=10,
        blank=True
    )


class RateCardHistory(models.Model):
    product = models.ForeignKey(
        Product,
        on_delete=models.CASCADE,
        related_name="history"
    )

    effective_date = models.CharField(
        max_length=50
    )

    snapshot = models.JSONField()

    created_at = models.DateTimeField(
        auto_now=True
    )

    class Meta:
        ordering = ["-created_at"]
        unique_together = ("product", "effective_date")

    def __str__(self):
        return f"{self.product.slug}:{self.effective_date}"