package io.project.UTM_bot.service;

import com.itextpdf.text.Document;
import com.itextpdf.text.Font;
import io.project.UTM_bot.config.BotConfig;
import org.json.JSONObject;
import org.springframework.stereotype.Component;
import org.telegram.telegrambots.bots.TelegramLongPollingBot;
import org.telegram.telegrambots.meta.api.methods.send.SendDocument;
import org.telegram.telegrambots.meta.api.methods.send.SendMessage;
import org.telegram.telegrambots.meta.api.objects.*;
import org.telegram.telegrambots.meta.api.objects.replykeyboard.InlineKeyboardMarkup;
import org.telegram.telegrambots.meta.api.objects.replykeyboard.buttons.*;

import java.io.*;
import java.io.File;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.List;

import com.itextpdf.text.*;
import com.itextpdf.text.pdf.PdfWriter;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

@Component
public class TelegramBot extends TelegramLongPollingBot {
    private final BotConfig config;
    private final Map<Long, UserState> userStates = new HashMap<>();

    public TelegramBot(BotConfig config) {
        this.config = config;
    }

    @Override
    public String getBotUsername() {
        return config.getBotName();
    }

    @Override
    public String getBotToken() {
        return config.getToken();
    }

    @Override
    public void onUpdateReceived(Update update) {
        if (update.hasMessage() && update.getMessage().hasText()) {
            Message message = update.getMessage();
            long chatId = message.getChatId();
            String text = message.getText().trim();

            if (text.equals("/schedule_for_the_day")) {
                UserState state = new UserState();
                state.setWeekly(false);
                userStates.put(chatId, state);
                sendMessage(chatId, "📘 Укажи, на каком ты курсе (1, 2, 3 и т.д.):");
                state.setStep(UserState.Step.AWAITING_COURSE);
                return;
            }

            if (text.equals("/weekly_schedule")) {
                UserState state = new UserState();
                state.setWeekly(true);
                userStates.put(chatId, state);
                sendMessage(chatId, "📘 Укажи, на каком ты курсе (1, 2, 3 и т.д.):");
                state.setStep(UserState.Step.AWAITING_COURSE);
                return;
            }

            UserState state = userStates.get(chatId);
            if (state != null) {
                handleUserStep(chatId, text, state);
            }
        }




        if (update.hasCallbackQuery()) {
            String callbackData = update.getCallbackQuery().getData();
            long chatId = update.getCallbackQuery().getMessage().getChatId();
            UserState state = userStates.get(chatId);

            switch (callbackData) {
                case "download_yes" -> sendFormatOptions(chatId);
                case "download_no" -> sendMessage(chatId, "🔁 Хорошо, выбери другую команду.");
                case "format_pdf" -> {
                    File file = generatePdfFromText(state.getLastWeeklySchedule());
                    sendFile(chatId, file);
                }
                case "format_txt" -> {
                    File file = generateTxtFromText(state.getLastWeeklySchedule());
                    sendFile(chatId, file);
                }
                case "format_xls" -> {
                    File file = generateExcelFromText(state.getLastWeeklySchedule());
                    sendFile(chatId, file);
                }
            }

            if (callbackData.startsWith("course_")) {
                String course = callbackData.replace("course_", "");
                state = userStates.computeIfAbsent(chatId, k -> new UserState());
                state.setCourse(course + " курс");
                state.setStep(UserState.Step.AWAITING_GROUP);
                sendMessage(chatId, "✍️ Введите номер группы:");
            } else if (callbackData.startsWith("day_")) {
                String day = callbackData.replace("day_", "");
                if (state != null && state.getStep() == UserState.Step.AWAITING_DAY_OF_WEEK) {
                    state.setDayOfWeek(day);
                    state.setStep(UserState.Step.AWAITING_WEEK_PARITY);
                    sendWeekParityOptions(chatId);
                }
            } else if (callbackData.startsWith("week_")) {
                String parity = callbackData.replace("week_", "");
                if (state != null && state.getStep() == UserState.Step.AWAITING_WEEK_PARITY) {
                    state.setWeekParity(parity.equals("even") ? "Четная" : "Нечетная");
                    state.setStep(UserState.Step.COMPLETED);
                    if (state.isWeekly()) {
                        sendWeeklyScheduleFromJson(state, chatId);
                    } else {
                        sendScheduleFromJson(state, chatId);
                    }
                }
            }
        }
    }

    private File generatePdfFromText(String text) {
        try {
            File file = File.createTempFile("schedule_", ".pdf");
            Document doc = new Document();
            PdfWriter.getInstance(doc, new FileOutputStream(file));
            doc.open();
            Font font = FontFactory.getFont(FontFactory.HELVETICA, 12, BaseColor.BLACK);
            for (String line : text.split("\n")) {
                doc.add(new Paragraph(line, font));
            }
            doc.close();
            return file;
        } catch (Exception e) {
            e.printStackTrace(); return null;
        }
    }

    private File generateTxtFromText(String text) {
        try {
            File file = File.createTempFile("schedule_", ".txt");
            Files.write(file.toPath(), text.getBytes(StandardCharsets.UTF_8));
            return file;
        } catch (IOException e) {
            e.printStackTrace(); return null;
        }
    }

    private File generateExcelFromText(String text) {
        try {
            String groupName = "unknown";

            for (String line : text.split("\n")) {
                if (line.contains("группы")) {
                    String[] parts = line.split("группы");
                    if (parts.length > 1) {
                        groupName = parts[1].trim().replaceAll("[^A-Za-z0-9]", "_");
                        break;
                    }
                }
            }

            String weekSuffix = "unknown";
            if (text.contains("Четная")) {
                weekSuffix = "even";
            } else if (text.contains("Нечетная")) {
                weekSuffix = "odd";
            }

            File file = File.createTempFile("schedule_" + groupName + "_" + weekSuffix + "_", ".xlsx");
            Workbook workbook = new XSSFWorkbook();
            Sheet sheet = workbook.createSheet("Schedule");

            Row header = sheet.createRow(0);
            header.createCell(0).setCellValue("Zilele");
            header.createCell(1).setCellValue("Orele");
            header.createCell(2).setCellValue("Informație");

            CellStyle wrapStyle = workbook.createCellStyle();
            wrapStyle.setWrapText(true);

            int rowNum = 1;
            String currentDay = null;
            String currentTime = null;
            String currentSubject = null;
            List<String> extraLines = new ArrayList<>();
            boolean dayPrinted = false;

            String[] lines = text.split("\n");
            for (int i = 0; i < lines.length; i++) {
                String line = lines[i].trim();

                if (line.startsWith("📆")) {
                    currentDay = line.replace("📆", "").replace("*", "").replace(":", "").trim();
                    dayPrinted = false;
                } else if (line.startsWith("⏰")) {
                    currentTime = line.replace("⏰", "").trim();
                } else if (line.startsWith("📚")) {
                    currentSubject = line.replace("📚", "").trim();
                    extraLines.clear();

                    // Собираем все строки после предмета, пока не встретим новую пару или день
                    for (int j = i + 1; j < lines.length; j++) {
                        String nextLine = lines[j].trim();
                        if (nextLine.startsWith("⏰") || nextLine.startsWith("📆") || nextLine.startsWith("📚") || nextLine.startsWith("---")) {
                            break;
                        }
                        extraLines.add(nextLine);
                    }

                    // Составляем строку информации
                    StringBuilder fullInfo = new StringBuilder(currentSubject);
                    for (String extra : extraLines) {
                        fullInfo.append("\n").append(extra);
                    }

                    Row row = sheet.createRow(rowNum++);
                    row.createCell(0).setCellValue(dayPrinted ? "" : currentDay);
                    row.createCell(1).setCellValue(currentTime);
                    Cell infoCell = row.createCell(2);
                    infoCell.setCellValue(fullInfo.toString().trim());
                    infoCell.setCellStyle(wrapStyle);
                    dayPrinted = true;
                }
            }

            // Автоматическая ширина столбцов
            for (int i = 0; i < 3; i++) {
                sheet.autoSizeColumn(i);
            }

            try (FileOutputStream output = new FileOutputStream(file)) {
                workbook.write(output);
            }
            workbook.close();
            return file;

        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }
    }



    private void sendFile(long chatId, File file) {
        if (file == null || !file.exists()) {
            sendMessage(chatId, "❌ Не удалось создать файл.");
            return;
        }
        SendDocument doc = new SendDocument();
        doc.setChatId(String.valueOf(chatId));
        doc.setDocument(new InputFile(file));
        doc.setCaption("📎 Вот ваше расписание");
        try {
            execute(doc);
        } catch (Exception e) {
            e.printStackTrace();
            sendMessage(chatId, "❌ Ошибка при отправке файла.");
        }
    }

    private void handleUserStep(long chatId, String text, UserState state) {
        switch (state.getStep()) {
            case AWAITING_COURSE:
                if (text.matches("[1-4]")) {
                    state.setCourse(text + " курс");
                    state.setStep(UserState.Step.AWAITING_GROUP);
                    sendMessage(chatId, "📚 Теперь укажи свою группу:");
                } else {
                    sendMessage(chatId, "❗ Введите корректный номер курса (1–4).");
                }
                break;

            case AWAITING_GROUP:
                state.setGroup(text.toUpperCase());
                if (state.isWeekly()) {
                    state.setStep(UserState.Step.AWAITING_WEEK_PARITY);
                    sendWeekParityOptions(chatId);
                } else {
                    state.setStep(UserState.Step.AWAITING_DAY_OF_WEEK);
                    sendDayOfWeekSelection(chatId);
                }
                break;

            case COMPLETED:
                sendMessage(chatId, "✅ Расписание уже получено. Введите /schedule_for_the_day или /weekly_schedule, чтобы начать заново.");
                break;
        }
    }

    private void sendDayOfWeekSelection(long chatId) {
        SendMessage message = new SendMessage(String.valueOf(chatId), "Выберите день недели:");
        InlineKeyboardMarkup markup = new InlineKeyboardMarkup();

        Map<String, String> days = new LinkedHashMap<>();
        days.put("Luni", "Luni");
        days.put("Marti", "Marți");
        days.put("Miercuri", "Miercuri");
        days.put("Joi", "Joi");
        days.put("Vineri", "Vineri");
        days.put("Sambata", "Sâmbătă");

        List<List<InlineKeyboardButton>> rows = new ArrayList<>();
        for (Map.Entry<String, String> entry : days.entrySet()) {
            InlineKeyboardButton button = new InlineKeyboardButton();
            button.setText(entry.getValue());
            button.setCallbackData("day_" + entry.getKey());
            rows.add(Collections.singletonList(button)); // каждая кнопка на отдельной строке
        }

        markup.setKeyboard(rows);
        message.setReplyMarkup(markup);
        executeMessage(message);
    }

    private void sendWeekParityOptions(long chatId) {
        SendMessage message = new SendMessage(String.valueOf(chatId), "Выберите неделю:");
        InlineKeyboardMarkup markup = new InlineKeyboardMarkup();
        markup.setKeyboard(List.of(
                List.of(createInlineButton("1 (Четная)", "week_even"), createInlineButton("2 (Нечетная)", "week_odd"))
        ));
        message.setReplyMarkup(markup);
        executeMessage(message);
    }

    private void sendScheduleFromJson(UserState state, long chatId) {
        sendScheduleBlock(state, chatId, state.getDayOfWeek());
    }

    private void sendWeeklyScheduleFromJson(UserState state, long chatId) {
        StringBuilder fullSchedule = new StringBuilder();
        fullSchedule.append("📓 *Полное недельное расписание* для группы ")
                .append(state.getGroup())
                .append(":\n\n");

        List<String> days = List.of("Luni", "Martí", "Miercuri", "Joi", "Vineri", "Sâmbătă");

        for (int i = 0; i < days.size(); i++) {
            String day = days.get(i);
            StringBuilder result = new StringBuilder("📆 *" + day + "*:\n");
            String courseNumber = state.getCourse().replaceAll("\\D+", "");
            String filePath = "src/main/resources/schedule_" + courseNumber + ".json";
            String group = state.getGroup();
            String week = state.getWeekParity();
            String mappedWeek = "Четная".equalsIgnoreCase(week) ? "Even" : "Odd";

            try {
                if (!Files.exists(Paths.get(filePath))) continue;

                String content = Files.readString(Paths.get(filePath));
                JSONObject json = new JSONObject(content);

                if (!json.has(group)) continue;
                JSONObject groupSchedule = json.getJSONObject(group);
                if (!groupSchedule.has(day)) continue;

                JSONObject daySchedule = groupSchedule.getJSONObject(day);
                List<String> timeSlots = new ArrayList<>();
                Iterator<String> keys = daySchedule.keys();
                while (keys.hasNext()) timeSlots.add(keys.next());

                timeSlots.sort(Comparator.comparing(s -> {
                    try {
                        return new SimpleDateFormat("HH.mm").parse(s.split("-")[0]);
                    } catch (Exception e) {
                        return new Date(0);
                    }
                }));

                boolean found = false;
                for (String timeSlot : timeSlots) {
                    Object entryObj = daySchedule.get(timeSlot);
                    if (entryObj instanceof JSONObject jsonEntry) {
                        String value = null;
                        if (jsonEntry.has(mappedWeek) && !jsonEntry.isNull(mappedWeek)) {
                            value = jsonEntry.getString(mappedWeek);
                        } else if (jsonEntry.has("Both") && !jsonEntry.isNull("Both")) {
                            value = jsonEntry.getString("Both");
                        }
                        if (value != null) {
                            found = true;
                            result.append("⏰ ").append(timeSlot).append("\n")
                                    .append("📚 ").append(value).append("\n\n");
                        }
                    }
                }

                if (found) {
                    fullSchedule.append(result.toString().trim());
                    if (i < days.size() - 1) {
                        fullSchedule.append("\n-----------------------------------------------\n");
                    } else {
                        fullSchedule.append("\n\n");
                    }
                }

            } catch (Exception e) {
                sendMessage(chatId, "⚠️ Ошибка при обработке расписания: " + e.getMessage());
                return;
            }
        }
        state.setLastWeeklySchedule(fullSchedule.toString().trim());
        sendMessage(chatId, state.getLastWeeklySchedule());
        sendDownloadOptions(chatId);
    }

    private void sendDownloadOptions(long chatId) {
        SendMessage message = new SendMessage(String.valueOf(chatId), "Хотите ли вы скачать расписание?");
        InlineKeyboardMarkup markup = new InlineKeyboardMarkup();
        markup.setKeyboard(List.of(
                List.of(createInlineButton("✅Да", "download_yes"), createInlineButton("❌Нет", "download_no"))
        ));
        message.setReplyMarkup(markup);
        executeMessage(message);
    }

    private void sendFormatOptions(long chatId) {
        SendMessage message = new SendMessage(String.valueOf(chatId), "Выберите формат для скачивания:");
        InlineKeyboardMarkup markup = new InlineKeyboardMarkup();
        markup.setKeyboard(List.of(
                List.of(
                        createInlineButton("PDF", "format_pdf"),
                        createInlineButton("TXT", "format_txt"),
                        createInlineButton("Excel", "format_xls")
                )
        ));
        message.setReplyMarkup(markup);
        executeMessage(message);
    }

    private void sendScheduleBlock(UserState state, long chatId, String day) {
        String courseNumber = state.getCourse().replaceAll("\\D+", "");
        String filePath = "src/main/resources/schedule_" + courseNumber + ".json";

        String group = state.getGroup();
        String week = state.getWeekParity();

        String mappedWeek = null;
        if ("Четная".equalsIgnoreCase(week)) mappedWeek = "Even";
        else if ("Нечетная".equalsIgnoreCase(week)) mappedWeek = "Odd";

        try {
            if (!Files.exists(Paths.get(filePath))) {
                sendMessage(chatId, "Файл не найден: " + filePath);
                return;
            }

            String content = Files.readString(Paths.get(filePath));
            JSONObject json = new JSONObject(content);

            if (!json.has(group)) {
                sendMessage(chatId, "Группа не найдена в расписании.");
                return;
            }

            JSONObject groupSchedule = json.getJSONObject(group);

            if (!groupSchedule.has(day)) return;

            JSONObject daySchedule = groupSchedule.getJSONObject(day);

            List<String> timeSlots = new ArrayList<>();
            Iterator<String> keys = daySchedule.keys();
            while (keys.hasNext()) timeSlots.add(keys.next());

            timeSlots.sort(Comparator.comparing(s -> {
                try {
                    return new SimpleDateFormat("HH.mm").parse(s.split("-")[0]);
                } catch (Exception e) {
                    return new Date(0);
                }
            }));

            StringBuilder result = new StringBuilder("📆 *" + day + "*:\n");
            boolean found = false;

            for (String timeSlot : timeSlots) {
                Object entryObj = daySchedule.get(timeSlot);
                if (entryObj instanceof JSONObject jsonEntry) {
                    String value = null;
                    if (mappedWeek != null && jsonEntry.has(mappedWeek) && !jsonEntry.isNull(mappedWeek)) {
                        value = jsonEntry.getString(mappedWeek);
                    } else if (jsonEntry.has("Both") && !jsonEntry.isNull("Both")) {
                        value = jsonEntry.getString("Both");
                    }
                    if (value != null) {
                        found = true;
                        result.append("📚  ").append(timeSlot).append("\n").append("⏰ ").append(value).append("\n\n");
                    }
                }
            }

            if (found) sendMessage(chatId, result.toString().trim());

        } catch (Exception e) {
            sendMessage(chatId, "⚠️ Ошибка при обработке расписания: " + e.getMessage());
        }
    }

    private void sendMessage(long chatId, String text) {
        SendMessage message = new SendMessage(String.valueOf(chatId), text);
        message.setParseMode("Markdown");
        executeMessage(message);
    }

    private void executeMessage(SendMessage message) {
        try {
            execute(message);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private InlineKeyboardButton createInlineButton(String text, String callbackData) {
        InlineKeyboardButton button = new InlineKeyboardButton();
        button.setText(text);
        button.setCallbackData(callbackData);
        return button;
    }

    private static class UserState {
        public String getLastWeeklySchedule() {
            return lastWeeklySchedule;
        }

        public void setLastWeeklySchedule(String lastWeeklySchedule) {
            this.lastWeeklySchedule = lastWeeklySchedule;
        }

        enum Step {
            AWAITING_COURSE,
            AWAITING_GROUP,
            AWAITING_DAY_OF_WEEK,
            AWAITING_WEEK_PARITY,
            COMPLETED
        }

        private Step step = Step.AWAITING_COURSE;
        private boolean weekly = false;
        private String course;
        private String group;
        private String dayOfWeek;
        private String weekParity;
        private String lastWeeklySchedule;

        public Step getStep() { return step; }
        public void setStep(Step step) { this.step = step; }
        public boolean isWeekly() { return weekly; }
        public void setWeekly(boolean weekly) { this.weekly = weekly; }
        public String getCourse() { return course; }
        public void setCourse(String course) { this.course = course; }
        public String getGroup() { return group; }
        public void setGroup(String group) { this.group = group; }
        public String getDayOfWeek() { return dayOfWeek; }
        public void setDayOfWeek(String dayOfWeek) { this.dayOfWeek = dayOfWeek; }
        public String getWeekParity() { return weekParity; }
        public void setWeekParity(String weekParity) { this.weekParity = weekParity; }
    }
}
