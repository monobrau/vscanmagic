namespace VScanMagic.ConnectSecure;

public sealed class ConnectSecureSuppressService(ConnectSecureClient client)
{
    public async Task<PatchOperationResult> SuppressAsync(
        SuppressVulnerabilityRequest request,
        CancellationToken ct = default)
    {
        if (request.CompanyId <= 0)
            throw new InvalidOperationException("Company id is required.");
        if (request.SolutionId <= 0)
            throw new InvalidOperationException("Solution id is required.");
        if (string.IsNullOrWhiteSpace(request.Reason))
            throw new InvalidOperationException("A suppression reason is required.");

        var data = new Dictionary<string, object?>
        {
            ["company_id"] = request.CompanyId,
            ["solution_id"] = request.SolutionId,
            ["reason"] = request.Reason.Trim(),
            ["suppress_comments"] = request.Comments.Trim(),
            ["problem_name"] = request.Product.Trim(),
            ["suppression_status"] = "Approved"
        };

        if (request.AssetId > 0)
        {
            data["asset_id"] = request.AssetId;
        }

        var response = await client.InvokeAuthenticatedAsync(
            HttpMethod.Post,
            "/w/asset/suppress_vulnerability",
            body: new Dictionary<string, object?> { ["data"] = data },
            ct: ct);

        var success = ConnectSecureJsonReader.GetBool(response, "status");
        var message = ConnectSecureJsonReader.GetString(response, "message");
        if (string.IsNullOrWhiteSpace(message))
            message = success ? "Suppression request accepted." : "ConnectSecure rejected the suppression request.";

        if (!success)
            throw new InvalidOperationException(message);

        return new PatchOperationResult(true, message);
    }
}
