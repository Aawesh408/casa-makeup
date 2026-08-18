<?php
// ============================================================
// CASA MAKE-UP - BOOKING HANDLER
// Receives bookings, saves to Excel, sends email notification
// ============================================================

header('Content-Type: application/json');
header('X-Content-Type-Options: nosniff');

// Only POST is allowed
if ($_SERVER['REQUEST_METHOD'] !== 'POST') {
    http_response_code(405);
    echo json_encode(['success' => false, 'message' => 'Method not allowed.']);
    exit;
}

require_once __DIR__ . '/config.php';

if (!file_exists(__DIR__ . '/vendor/autoload.php')) {
    header('Content-Type: application/json');
    http_response_code(500);
    echo json_encode(['success' => false, 'message' => 'Booking system is temporarily unavailable. Please contact us directly.']);
    exit;
}
require_once __DIR__ . '/vendor/autoload.php';

date_default_timezone_set('Asia/Kolkata');

// ---- Read input ----

$input = json_decode(file_get_contents('php://input'), true);
if (!is_array($input) || empty($input)) {
    $input = $_POST;
}

$name         = isset($input['name']) ? trim($input['name']) : '';
$phone        = isset($input['phone']) ? trim($input['phone']) : '';
$email        = isset($input['email']) ? trim($input['email']) : '';
$service      = isset($input['service']) ? trim($input['service']) : '';
$date         = isset($input['date']) ? trim($input['date']) : '';
$time         = isset($input['time']) ? trim($input['time']) : '';
$address      = isset($input['address']) ? trim($input['address']) : '';
$instructions = isset($input['instructions']) ? trim($input['instructions']) : '';

if (empty($instructions) && isset($input['message'])) {
    $instructions = trim($input['message']);
}

// Limit lengths
$name         = mb_substr($name, 0, 100);
$phone        = mb_substr($phone, 0, 15);
$email        = mb_substr($email, 0, 100);
$service      = mb_substr($service, 0, 50);
$date         = mb_substr($date, 0, 10);
$time         = mb_substr($time, 0, 10);
$address      = mb_substr($address, 0, 200);
$instructions = mb_substr($instructions, 0, 500);

// ---- Validation ----

function sendError($msg) {
    http_response_code(422);
    echo json_encode(['success' => false, 'message' => $msg]);
    exit;
}

if (empty($name)) {
    sendError('Please enter your full name.');
}

// Phone: strip spaces, dashes, parentheses, dots
$phone = preg_replace('/[\s\-\(\)\.]/', '', $phone);

// Remove leading zeros
if (strpos($phone, '0') === 0) {
    $phone = ltrim($phone, '0');
}

// Strip country code
if (preg_match('/^91\d{10}$/', $phone) && strlen($phone) === 12) {
    $phone = substr($phone, 2);
}
if (strpos($phone, '+91') === 0) {
    $phone = substr($phone, 3);
}

if (!preg_match('/^[6-9][0-9]{9}$/', $phone)) {
    sendError('Please enter a valid 10-digit Indian mobile number.');
}

if (!empty($email) && !filter_var($email, FILTER_VALIDATE_EMAIL)) {
    sendError('Please enter a valid email address.');
}

if (empty($service)) {
    sendError('Please select a service.');
}

$validServices = [
    'Bridal Makeup', 'Party Makeup', 'Reception Makeup',
    'Hair Styling', 'Hair Color', 'Facial',
    'Nail Art & Manicure', 'Full Body Wax', 'Other'
];
if (!in_array($service, $validServices)) {
    sendError('Please select a valid service.');
}

if (empty($date) || !preg_match('/^\d{4}-\d{2}-\d{2}$/', $date)) {
    sendError('Please select a valid date.');
}

$dp = explode('-', $date);
if (!checkdate((int)$dp[1], (int)$dp[2], (int)$dp[0])) {
    sendError('Please enter a valid calendar date.');
}

$bookingDate = new DateTime($date, new DateTimeZone('Asia/Kolkata'));
$today       = new DateTime('today', new DateTimeZone('Asia/Kolkata'));
if ($bookingDate < $today) {
    sendError('Date cannot be in the past. Please select today or a future date.');
}

if (empty($time) || !preg_match('/^\d{2}:\d{2}$/', $time)) {
    sendError('Please select a valid time.');
}

$tp = explode(':', $time);
$minutes = (int)$tp[0] * 60 + (int)$tp[1];
if ($minutes < 600 || $minutes > 1200) {
    sendError('Please select a time between 10:00 AM and 8:00 PM.');
}

// ---- Server-generated values ----

$bookingId = 'CM' . date('Ymd') . strtoupper(substr(md5(uniqid(mt_rand(), true)), 0, 4));
$timestamp = date('Y-m-d H:i:s');

// Prefix dangerous characters to prevent Excel formula injection
function excelSafe($val) {
    if ($val === '' || $val === null) {
        return $val;
    }
    if (in_array(substr($val, 0, 1), ['=', '+', '-', '@'], true)) {
        return "'" . $val;
    }
    return $val;
}

// ---- Save to Excel ----

$xlsxFile = BOOKINGS_DIR . '/Casa_Makeup_Bookings.xlsx';
$lockFile = BOOKINGS_DIR . '/.booking.lock';

try {
    if (!is_dir(BOOKINGS_DIR)) {
        mkdir(BOOKINGS_DIR, 0755, true);
    }

    // Delete stale Excel lock files (~$filename.xlsx) left by Microsoft Office
    $staleLocks = glob(BOOKINGS_DIR . '/~$*.xlsx');
    if (is_array($staleLocks)) {
        foreach ($staleLocks as $staleFile) {
            @unlink($staleFile);
        }
    }

    // File lock to prevent concurrent write corruption
    $lockFp = fopen($lockFile, 'c');
    if ($lockFp === false) {
        throw new RuntimeException('Cannot create lock file.');
    }
    flock($lockFp, LOCK_EX);

    if (file_exists($xlsxFile)) {
        $reader   = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
        $reader->setReadDataOnly(false);
        $spreadsheet = $reader->load($xlsxFile);
    } else {
        $spreadsheet = new \PhpOffice\PhpSpreadsheet\Spreadsheet();
    }

    // Get or create the Bookings sheet
    $sheet = $spreadsheet->getSheetByName('Bookings');
    if ($sheet === null) {
        $sheet = $spreadsheet->getActiveSheet();
        $sheet->setTitle('Bookings');
    }

    // Check if header row already exists
    $hasHeader = ($sheet->getCell('A1')->getValue() !== null);
    $nextRow   = $sheet->getHighestRow() + 1;

    if (!$hasHeader) {
        $nextRow = 2;

        $headers = [
            'Booking ID', 'Customer Name', 'Phone', 'Email', 'Service',
            'Preferred Date', 'Preferred Time', 'Address', 'Special Instructions', 'Booking Timestamp'
        ];

        foreach ($headers as $col => $text) {
            $sheet->setCellValueByColumnAndRow($col + 1, 1, $text);
        }

        // Header style
        $headerStyle = [
            'font'      => ['bold' => true, 'color' => ['rgb' => 'FFFFFF'], 'name' => 'Calibri', 'size' => 11],
            'fill'      => ['fillType' => 'solid', 'color' => ['rgb' => '2D2A26']],
            'alignment' => ['horizontal' => 'center', 'vertical' => 'center'],
            'borders'   => ['bottom' => ['borderStyle' => 'thin', 'color' => ['rgb' => 'B8956A']]]
        ];
        $sheet->getStyle('A1:J1')->applyFromArray($headerStyle);
        $sheet->getRowDimension(1)->setRowHeight(30);

        // Column widths
        $widths = ['A' => 18, 'B' => 25, 'C' => 18, 'D' => 30, 'E' => 25,
                   'F' => 18, 'G' => 18, 'H' => 40, 'I' => 45, 'J' => 24];
        foreach ($widths as $col => $w) {
            $sheet->getColumnDimension($col)->setWidth($w);
        }
    }

    // Write the new booking row
    $rowData = [
        $bookingId,
        excelSafe($name),
        excelSafe($phone),
        excelSafe($email),
        excelSafe($service),
        $date,
        $time,
        excelSafe($address),
        excelSafe($instructions),
        $timestamp
    ];

    foreach ($rowData as $col => $value) {
        $sheet->setCellValueByColumnAndRow($col + 1, $nextRow, $value);
    }

    // Alternating row background
    $bgColor = ($nextRow % 2 === 0) ? 'F5F0E8' : 'FFFFFF';

    $dataStyle = [
        'font'      => ['name' => 'Calibri', 'size' => 10, 'color' => ['rgb' => '2D2A26']],
        'fill'      => ['fillType' => 'solid', 'color' => ['rgb' => $bgColor]],
        'alignment' => ['vertical' => 'center'],
        'borders'   => ['allBorders' => ['borderStyle' => 'thin', 'color' => ['rgb' => 'E0D8CE']]]
    ];
    $sheet->getStyle('A' . $nextRow . ':J' . $nextRow)->applyFromArray($dataStyle);

    // Wrap text for Address (H) and Special Instructions (I)
    $sheet->getStyle('H' . $nextRow)->getAlignment()->setWrapText(true);
    $sheet->getStyle('I' . $nextRow)->getAlignment()->setWrapText(true);

    // Phone as text to avoid scientific notation
    $sheet->getStyle('C' . $nextRow)->getNumberFormat()->setFormatCode('@');

    $sheet->getRowDimension($nextRow)->setRowHeight(22);

    // AutoFilter and freeze pane (safe to reapply)
    $sheet->setAutoFilter('A1:J1');
    $sheet->freezePane('A2');

    // Save to temp file first, then replace — prevents corruption if file is locked
    $tempXlsx = tempnam(BOOKINGS_DIR, 'booking_');
    if ($tempXlsx === false) {
        throw new RuntimeException('Cannot create temp file.');
    }

    // Suppress warnings during write so they don't corrupt JSON output
    $prevErrors = error_reporting();
    error_reporting($prevErrors & ~E_WARNING & ~E_NOTICE);
    $writer = new \PhpOffice\PhpSpreadsheet\Writer\Xlsx($spreadsheet);
    $writer->save($tempXlsx);
    error_reporting($prevErrors);

    // Replace original with the temp file
    if (!rename($tempXlsx, $xlsxFile)) {
        // If rename fails (e.g. Windows file lock), try copy + delete
        if (copy($tempXlsx, $xlsxFile)) {
            @unlink($tempXlsx);
        } else {
            @unlink($tempXlsx);
            throw new RuntimeException('Cannot save booking file. The file may be open in Excel. Please close it and try again.');
        }
    }

    // Release lock
    flock($lockFp, LOCK_UN);
    fclose($lockFp);
    @unlink($lockFile);

} catch (\Exception $e) {
    if (isset($lockFp) && is_resource($lockFp)) {
        flock($lockFp, LOCK_UN);
        fclose($lockFp);
    }
    error_log('Casa Makeup Excel Error [' . $timestamp . ']: ' . $e->getMessage() . ' | Booking ID: ' . $bookingId);
    http_response_code(500);
    echo json_encode(['success' => false, 'message' => 'We could not save your booking right now. Please try again or contact us directly.']);
    exit;
}

// ---- Send Email Notification ----

$emailSubject = 'New Booking - ' . BUSINESS_SHORT_NAME;

$emailBody = "Casa Make-up - New Booking Received\n"
    . "====================================\n\n"
    . "Booking ID:           $bookingId\n"
    . "Customer Name:        $name\n"
    . "Phone:                $phone\n"
    . "Email:                " . ($email ?: 'Not provided') . "\n"
    . "Service:              $service\n"
    . "Preferred Date:       $date\n"
    . "Preferred Time:       $time\n"
    . "Address:              " . ($address ?: 'Not provided') . "\n"
    . "Special Instructions: " . ($instructions ?: 'None') . "\n"
    . "Booking Timestamp:    $timestamp\n\n"
    . "---\n"
    . "This is an automated message from the Casa Make-up website.\n";

try {
    $mail = new \PHPMailer\PHPMailer\PHPMailer(true);
    $mail->isSMTP();
    $mail->Host       = SMTP_HOST;
    $mail->SMTPAuth   = true;
    $mail->Username   = SMTP_EMAIL;
    $mail->Password   = SMTP_PASSWORD;
    $mail->SMTPSecure = \PHPMailer\PHPMailer\PHPMailer::ENCRYPTION_STARTTLS;
    $mail->Port       = SMTP_PORT;
    $mail->CharSet    = 'UTF-8';

    $mail->setFrom(SMTP_EMAIL, SMTP_FROM_NAME);
    $mail->addAddress(MAIL_TO, MAIL_TO_NAME);
    $mail->isHTML(false);

    if (!empty($email)) {
        $mail->addReplyTo($email, $name);
    }

    $mail->Subject = $emailSubject;
    $mail->Body    = $emailBody;
    $mail->send();

} catch (\Exception $e) {
    // Email failed but booking is saved - log error and continue
    error_log('Casa Makeup Email Error [' . $timestamp . ']: ' . $e->getMessage() . ' | Booking ID: ' . $bookingId);
}

// ---- Return success response ----

echo json_encode([
    'success'   => true,
    'message'   => 'Your booking request has been successfully received. Our team will contact you shortly.',
    'bookingId' => $bookingId
]);
exit;
