<?php
// ============================================================
// CASA MAKE-UP - CONFIGURATION
// Change these values after deployment
// ============================================================

// Business Information
define('BUSINESS_NAME', 'Casa Make-up by Sonali Shukla');
define('BUSINESS_SHORT_NAME', 'Casa Make-up');
define('BUSINESS_TAGLINE', 'Premium Bridal & Party Makeup Artist in Lucknow');

// Contact Information
define('OWNER_PHONE', '9598848212');
define('OWNER_PHONE_DISPLAY', '+91 9598848212');
define('OWNER_EMAIL', 'aaweshtiwari6388@gmail.com');
define('OWNER_ADDRESS', 'Gonda, Uttar Pradesh 271001');

// Working Hours
define('WORKING_HOURS_START', '10:00');
define('WORKING_HOURS_END', '20:00');
define('WORKING_HOURS_TEXT', 'Monday - Sunday, 10:00 AM - 8:00 PM');

// SMTP Configuration (Gmail App Password)
define('SMTP_HOST', 'smtp.gmail.com');
define('SMTP_PORT', 587);
define('SMTP_EMAIL', 'aaweshtiwari6388@gmail.com');           // Change this
define('SMTP_PASSWORD', 'mjsuijxanqioojum');     // Change this - use App Password, NOT normal password
define('SMTP_FROM_NAME', 'Casa Make-up Website');

// Email Recipient
define('MAIL_TO', 'aaweshtiwari6388@gmail.com');
define('MAIL_TO_NAME', 'Sonali Shukla');

// File Paths
define('BOOKINGS_DIR', __DIR__ . '/admin');
define('BOOKINGS_CSV', BOOKINGS_DIR . '/bookings.csv');

// WhatsApp Link
define('WHATSAPP_LINK', 'https://wa.me/91' . OWNER_PHONE);

// Business Stats (easily editable)
define('STAT_RATING', '4.9');
define('STAT_CLIENTS', '2500+');
define('STAT_BRIDAL_LOOKS', '500+');
define('STAT_EXPERIENCE', '2+');
