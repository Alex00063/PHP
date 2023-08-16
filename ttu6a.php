<?php
include '../reports.php';

		$to='a.nazarov.2020@mail.ru'; // Получатели
		$thm='Отчет о поездках'; // Тема псьма
		$path=Reports::rep6ATTU();
		$fp = fopen($path,"r"); 
		if (!$fp) 
		{ 
			print "Файл $path не может быть прочитан"; 
			exit(); 
		} 
		$file = fread($fp, filesize($path)); 
		fclose($fp); 
		$boundary = "--".md5(uniqid(time())); // генерируем разделитель 
		$headers .= "MIME-Version: 1.0\n"; 
		$headers .="Content-Type: multipart/mixed; boundary=\"$boundary\"\n"; 
		$multipart .= "--$boundary\n"; 
		$kod = 'koi8-r'; // или $kod = 'windows-1251'; 
		$multipart .= "Content-Type: text/html; charset=$kod\n"; 
		$multipart .= "Content-Transfer-Encoding: Quot-Printed\n\n"; 
		$multipart .= "$html\n\n"; 
		$message_part = "--$boundary\n"; 
		$message_part .= "Content-Type: application/octet-stream\n"; 
		$message_part .= "Content-Transfer-Encoding: base64\n"; 
		$message_part .= "Content-Disposition: attachment; filename = \"".basename($path)."\"\n\n"; // Вложение
		$message_part .= chunk_split(base64_encode($file))."\n"; 
		$multipart .= $message_part."--$boundary--\n"; 
		if(!mail($to, $thm, $multipart, $headers)) 
		{ 
			echo "К сожалению, письмо не отправлено"; 
			exit; 
		}
		unlink($path);
		echo "<script>window.close()</script>";
?>