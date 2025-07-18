def analyze_import_data(incoming_data):
    """
    Compares incoming automation data with existing DB entries.
    Returns lists of automations to create, update (with diffs), and delete.
    """
    existing_automations = {a.air_id: a for a in Automation.objects.all()}
    incoming_air_ids = set()
    to_create = []
    to_update = []
    to_delete = []

    for item in incoming_data:
        air_id = item.get("air_id")
        if not air_id:
            continue  # skip if air_id is missing
        incoming_air_ids.add(air_id)
        existing = existing_automations.get(air_id)

        if existing:
            changes = {}
            for field, new_value in item.items():
                if hasattr(existing, field):
                    old_value = getattr(existing, field)
                    if str(old_value) != str(new_value):
                        changes[field] = {"old": old_value, "new": new_value}
            if changes:
                to_update.append({
                    "air_id": air_id,
                    "changes": changes,
                    "updated_fields": list(changes.keys())
                })
        else:
            to_create.append(item)

    # Anything in DB but not in the incoming data => delete
    to_delete = [
        air_id for air_id in existing_automations.keys()
        if air_id not in incoming_air_ids
    ]

    return to_create, to_update, to_delete


@action(detail=False, methods=["post"])
def handleImport(self, request):
    """
    POST /automations/handleImport/
    Accepts a list of automation dicts, compares them to DB,
    creates new ones, updates changed ones, and deletes stale ones.
    """
    data = request.data
    if not isinstance(data, list):
        return Response({"error": "Expected a list of automations."}, status=status.HTTP_400_BAD_REQUEST)

    to_create, to_update, to_delete = analyze_import_data(data)

    # Perform bulk create
    created_objs = Automation.objects.bulk_create([
        Automation(**item) for item in to_create
    ])

    # Perform updates
    updated_objs = []
    for update in to_update:
        air_id = update["air_id"]
        obj = Automation.objects.get(air_id=air_id)
        for field in update["updated_fields"]:
            setattr(obj, field, update["changes"][field]["new"])
        obj.save()
        updated_objs.append(obj)

    # Perform deletions
    Automation.objects.filter(air_id__in=to_delete).delete()

    return Response({
        "created": [obj.air_id for obj in created_objs],
        "updated": [obj.air_id for obj in updated_objs],
        "deleted": to_delete,
        "summary": {
            "created": len(created_objs),
            "updated": len(updated_objs),
            "deleted": len(to_delete),
        }
    }, status=status.HTTP_200_OK)
