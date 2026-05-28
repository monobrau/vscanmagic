using System.Text.Json;

namespace VScanMagic.ConnectSecure;

public sealed class ConnectSecureSuppressService(ConnectSecureClient client)
{
    public async Task<PatchOperationResult> SuppressAsync(
        SuppressVulnerabilityRequest request,
        CancellationToken ct = default)
    {
        if (request.CompanyId <= 0)
            throw new InvalidOperationException("Company id is required.");
        if (request.SolutionId <= 0 && request.ProblemId <= 0)
            throw new InvalidOperationException("Solution id or problem id is required.");
        if (string.IsNullOrWhiteSpace(request.Reason))
            throw new InvalidOperationException("A suppression reason is required.");

        var data = new Dictionary<string, object?>
        {
            ["company_id"] = request.CompanyId,
            ["reason"] = request.Reason.Trim(),
            ["suppress_comments"] = request.Comments.Trim(),
            ["suppression_status"] = "Approved"
        };

        if (request.ProblemId > 0)
        {
            if (string.IsNullOrWhiteSpace(request.Product))
                throw new InvalidOperationException("Problem name is required for CVE suppression.");

            data["problem_id"] = request.ProblemId;
            data["problem_name"] = request.Product.Trim();
        }
        else
        {
            data["solution_id"] = request.SolutionId;
            data["problem_name"] = request.Product.Trim();
        }

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

    public async Task<PatchOperationResult> UnsuppressAsync(int suppressRecordId, CancellationToken ct = default)
    {
        if (suppressRecordId <= 0)
            throw new ArgumentOutOfRangeException(nameof(suppressRecordId), "Suppress record id is required.");

        var response = await client.InvokeAuthenticatedAsync(
            HttpMethod.Delete,
            $"/d/asset/suppress_vulnerability/{suppressRecordId}",
            ct: ct);

        var success = ConnectSecureJsonReader.GetBool(response, "status");
        var message = ConnectSecureJsonReader.GetString(response, "message");
        if (string.IsNullOrWhiteSpace(message))
            message = success ? "Suppression removed." : "ConnectSecure rejected the unsuppress request.";

        if (!success)
            throw new InvalidOperationException(message);

        return new PatchOperationResult(true, message);
    }

    public async Task<int?> FindSuppressRecordIdAsync(
        int companyId,
        int? problemId = null,
        int? solutionId = null,
        CancellationToken ct = default)
    {
        if (companyId <= 0)
            return null;

        if (problemId is > 0)
        {
            var byProblem = await FetchSuppressRecordsAsync(
                companyId,
                $"company_id={companyId} and problem_id={problemId.Value}",
                ct);
            var id = FirstSuppressRecordId(byProblem);
            if (id is > 0)
                return id;
        }

        if (solutionId is > 0)
        {
            var bySolution = await FetchSuppressRecordsAsync(
                companyId,
                $"company_id={companyId} and solution_id={solutionId.Value}",
                ct);
            return FirstSuppressRecordId(bySolution);
        }

        return null;
    }

    private static int? FirstSuppressRecordId(IEnumerable<JsonElement> rows)
    {
        foreach (var row in rows)
        {
            var id = ConnectSecureJsonReader.GetInt(row, "id") ?? 0;
            if (id > 0)
                return id;
        }

        return null;
    }

    private async Task<List<JsonElement>> FetchSuppressRecordsAsync(
        int companyId,
        string condition,
        CancellationToken ct)
    {
        var response = await client.InvokeAuthenticatedAsync(
            HttpMethod.Get,
            "/r/asset/suppress_vulnerability",
            new Dictionary<string, string>
            {
                ["condition"] = condition,
                ["limit"] = "10",
                ["skip"] = "0"
            },
            ct: ct);

        return ConnectSecureJsonReader.ExtractDataArray(response);
    }
}
