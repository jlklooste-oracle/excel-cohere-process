# Generative AI Service DP TypeScript API demo

**How to run \*.ts on your laptop :**

## 0. Set up
```
Follow link below to generate a config and session token for your oci user and upload it to oci_console.
https://docs.oracle.com/en-us/iaas/Content/API/Concepts/sdk_authentication_methods.htm

After completion, you should have the following values (also can be obtained from "View Configuration file"): 

a. CONFIG_PROFILE: the name of your config profile;
b. compartmentId: id of your tenancy.

```
## 1. Install TypeScript SDK
```
npm config set registry https://artifactory.oci.oraclecorp.com/api/npm/global-dev-npm
npm install oci-sdk@2.70.1-preview-1694034073
```

## 2. Make sure you are in the root directory of this example:
```
cd ~/genai-demo/genai-data-plane-ts-sample
```

## 3. Update the session token
Update the following field with the values from above. An example is below
```
compartmentId="ocid1.tenancy.oc1..."
CONFIG_PROFILE="genaiusers"
```

## 4. Kick off different shell sample script directly
```
tsc ./<script_name>.ts
node ./<script_name>.js
```

## 5. Example Results

./generate_text_example.ts
```
{
  "id": "c86e51e3-65b8-4a79-8b42-5d3c8742e0ed",
  "generatedTexts": [
    [
      {
        "id": "c86e51e3-65b8-4a79-8b42-5d3c8742e0ed",
        "text": " The Earth is the only planet in the solar system that is known to have liquid water on its surface"
      }
    ]
  ],
  "timeCreated": "2023-09-13T05:04:46.832Z",
  "modelId": "cohere.command",
  "modelVersion": "14.2"
}
```

./summarize_text_example.ts
```
{
  "id": "10364f8a-e553-42e8-87ca-285b0ae5d501",
  "summary": "Semiconducting qubits could enable larger and more complex quantum systems, says Google researcher Maharaja jan dark. Such qubits could also aid in quantum error correction, and address scalability and stability challenges. Google has been developing superconducting qubits, while Microsoft has focused on topological qubits which may be easier to manufacture and more reliable to operate.",
  "modelId": "cohere.command",
  "modelVersion": "14.2"
}
```

./embed_text_example.ts
```
{
  "id": "1ae132d3-1598-41b1-9a4a-5a56513c23a5",
  "embeddings": [
    [
      -2.5859375,
      0.27807617,
      ...
      -0.11608887,
      -1.4052734,
      -2.1484375
    ],
    [
      -2.203125,
      -1.5107422,
      -0.4790039,
      ...
      0.05706787,
      -0.80322266
    ]
  ],
  "modelId": "ocid1.generativeaimodel.oc1.us-chicago-1.amaaaaaapi24rzaauhcci2xuk7263dbon4xnkjku5ifv6i45nzbz6rrn4joa",
  "modelVersion": "2.0"
}
```