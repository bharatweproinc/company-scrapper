<?php
// Allow CORS
header("Access-Control-Allow-Origin: *");
header("Access-Control-Allow-Methods: GET, POST, OPTIONS");
header("Access-Control-Allow-Headers: Content-Type, Authorization");

if ($_SERVER['REQUEST_METHOD'] === 'OPTIONS') {
    http_response_code(200);
    exit;
}

header('Content-Type: application/json');

// Read POST body (JSON or form-data)
$input = file_get_contents("php://input");
$postData = json_decode($input, true);

// If not JSON, fallback to $_POST (form-data or x-www-form-urlencoded)
if (json_last_error() !== JSON_ERROR_NONE) {
    $postData = $_GET;
}

// Extract fields
$totalRecord = $postData['totalRecord'] ?? null;
$switchTo = $postData['switchTo'] ?? null;

if (!$totalRecord || !$switchTo) {
    echo json_encode([
        'success' => false,
        'message' => 'Missing required field: totalRecord or switchTo',
        'type' => "Gmail"
    ]);
    exit;
}

// Determine PowerShell script path
switch ($switchTo) {
    case "tenants-companies":
        $path = 'tenant-companies-list.ps1';
        break;
    default:
        $path = null;
}

$results = [];

if (!$path) {
    echo json_encode([
        'success' => false,
        'message' => 'Invalid type. Use reply | send | send_and_reply'
    ]);
    exit;
}

$scriptPath = __DIR__ . '/' . $path;
if (!file_exists($scriptPath)) {
    $results = [
        'path' => $path,
        'success' => false,
        'message' => 'Script not found'
    ];
}

// Build PowerShell command
$OnlyReplyTo = isset($postData['only_reply_to']) && $postData['only_reply_to'] 
    ? json_encode($postData['only_reply_to']) 
    : "[]";

$cmd = "powershell -ExecutionPolicy Bypass -File \"$scriptPath\""
     . " -totalRecord \"$totalRecord\"";

// $cmd = "powershell -ExecutionPolicy Bypass -File \"$scriptPath\" -Email \"$email\" -ReplyTo \"$ReplyTo\" -Subject \"$Subject\" -Recipient \"$Recipient\" -MessageBody \"$MessageBody\" -OnlyReplyTo \"$OnlyReplyTo\" -CheckForEmail \"$CheckForEmail\"";

$output = [];
$returnVar = 0;
exec($cmd . " 2>&1", $output, $returnVar);

// Join all output
$outputString = trim(implode("\n", $output));

// --- Extract only the JSON block ---
$jsonString = null;
$jsonStart  = strpos($outputString, '{');
$jsonEnd    = strrpos($outputString, '}');

if ($jsonStart !== false && $jsonEnd !== false) {
    $jsonString = substr($outputString, $jsonStart, $jsonEnd - $jsonStart + 1);
}

$jsonString = trim($jsonString, "\xEF\xBB\xBF");
$jsonString = mb_convert_encoding($jsonString, 'UTF-8', 'UTF-8');

$data = $jsonString ? json_decode($jsonString, true) : null;

if (json_last_error() === JSON_ERROR_NONE && is_array($data)) {
    $results = [
        'path' => $path,
        'success' => $data['success'] ?? false,
        'message' => $data['message'] ?? '',
        'data' => $data['data'] ?? null,
    ];
} else {
    $results = [
        'path' => $path,
        'success' => false,
        'message' => 'Invalid JSON from script',
    ];
}

// Final combined response
echo json_encode([
    'success' => true,
    'results' => $results
], JSON_PRETTY_PRINT);
exit;
