@action(detail=False, methods=['post'])
def bulk_sync(self, request):
    """
    Sync imported automation records:
    - Create new records.
    - Update existing ones only if fields differ.
    - Delete automations not present in the import.
    """
    incoming_data = request.data
    if not isinstance(incoming_data, list):
        return Response({'error': 'Expected a list of automation objects'}, status=status.HTTP_400_BAD_REQUEST)

    incoming_air_ids = set()
    created, updated, skipped = 0, 0, 0
    existing_automations = {a.air_id: a for a in Automation.objects.all()}
    updated_automations = []

    for item in incoming_data:
        air_id = item.get('air_id')
        if not air_id:
            continue  # skip invalid record

        incoming_air_ids.add(air_id)
        existing = existing_automations.get(air_id)

        if existing:
            # Only update if any fields differ
            serializer = AutomationSerializer(existing, data=item)
            if serializer.is_valid():
                if any(getattr(existing, field) != serializer.validated_data.get(field) for field in serializer.validated_data):
                    serializer.save()
                    updated += 1
                else:
                    skipped += 1
            else:
                return Response(serializer.errors, status=status.HTTP_400_BAD_REQUEST)
        else:
            serializer = AutomationSerializer(data=item)
            if serializer.is_valid():
                serializer.save()
                created += 1
            else:
                return Response(serializer.errors, status=status.HTTP_400_BAD_REQUEST)

    # Delete automations not in import
    to_delete = Automation.objects.exclude(air_id__in=incoming_air_ids)
    deleted_count, _ = to_delete.delete()

    return Response({
        'created': created,
        'updated': updated,
        'skipped': skipped,
        'deleted': deleted_count
    }, status=status.HTTP_200_OK)


const handleImport = async () => {
  setIsProcessing(true);
  setStep(3);

  const allRecords = [...importPreview.newRecords, ...importPreview.updateRecords];

  try {
    const response = await fetch('/api/automations/bulk_sync/', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify(allRecords),
    });

    const result = await response.json();
    if (response.ok) {
      alert(`Sync complete: ${result.created} created, ${result.updated} updated, ${result.skipped} unchanged, ${result.deleted} deleted`);
    } else {
      console.error(result);
      alert('Sync failed. Check console for details.');
    }

    if (onImport) await onImport();
    setIsProcessing(false);
    setTimeout(() => onClose(), 2000);
  } catch (error) {
    console.error('Import error:', error);
    alert(`Failed to import data: ${error.message}`);
    setIsProcessing(false);
  }
};



const response = await fetch(`${BACKEND_URL}/api/automations/bulk_sync/`, {
  method: 'POST',
  headers: {
    'Content-Type': 'application/json',
  },
  body: JSON.stringify(allRecords),
});
