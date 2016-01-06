(ns clexcel.core
  (:require [clj-time.core :as t]
            [clj-time.coerce :as tc])
  (:gen-class))

(use 'dk.ative.docjure.spreadsheet)

(defn load-diary
  "Load a diary into a map"
  []
  (->> (load-workbook "Diary.xls")
       (select-sheet "2015")
       (select-columns
        {:B :datum :C :von :D :bis :E :dauer :F :projekt :G :task :H :beschreibung :I :ueberstunden})))

(def tz-here (t/time-zone-for-id "Europe/Vienna"))

(defn timezoned
  [date]
  (when-not (nil? date) (t/to-time-zone (tc/to-date-time date) tz-here)))

(defn month-of
  "Returns the month of the datum cell of a row"
  [entry]
  (t/month (tc/to-local-date (timezoned (:datum entry)))))

(defn load-month
  [m]
  (filter (fn [entry] (= (month-of entry) m))
          (rest (load-diary))))

(defn map-values-to-vec
  [e]
  (row-vec [:datum :von :bis :dauer :projekt :task :beschreibung :ueberstunden] e))

(defn fix-t
  [time]
  (tc/to-date (t/plus (timezoned time) (t/days 1))))

(defn fix-time-von
  [entry]
  (update-in entry [:von] fix-t))

(defn fix-time-bis
  [entry]
  (update-in entry [:bis] fix-t))

(defn fix-times
  [raw]
  (map fix-time-bis (map fix-time-von raw)))

(defn prepare-for-excel
  "takes a vec of maps and prepares it for an excel sheet"
  [raw]
  (concat [["Mitarbeiter:" "Bernhard Sillipp" "" "Zeitraum:" "Mai 2015"]
           ["Datum" "Von" "Bis" "Dauer" "Projekt" "Task" "Beschreibung" "Überstunden"]]
        (map map-values-to-vec (fix-times raw))))

(defn format-header
  [wb header-rows]
  (doall
   (map #(set-row-style! % (create-cell-style! wb {:font {:bold true}})) header-rows)))

(defn col-seq
  "Returns a sequence of all cells in column col of the sheet"
  [col sheet]
  (seq (map #(.getCell % col) (row-seq sheet))))

(defn format-cell
  [cell format]
  (when-not (nil? cell)
    (apply-date-format! cell format)))

(defn format-col
  [sheet col fmt]
  (doseq [cell (col-seq col sheet)]
    (format-cell cell fmt)))

(defn format-cols
  [sheet]
  (do
    (format-col sheet 0 "dd.MM.yyyy")
    (format-col sheet 1 "hh:mm")
    (format-col sheet 2 "hh:mm")
    (format-col sheet 3 "0.00")
    (format-col sheet 4 "@")
    (format-col sheet 5 "@")
    (format-col sheet 6 "@")
    (format-col sheet 7 "0.00")))

(defn format-col-size
  [sheet]
  (doseq [col (range 8)]
    (.autoSizeColumn sheet col)))

(defn datum-of
  [row]
  (when-not (empty? row)
    (if (= 0 (.getCellType (first row)))
      (.getDateCellValue (first row)))))

(defn different-day?
  ([row1]
   false)
  ([row1 row2]
   (not= (datum-of row1) (datum-of row2))))

(defn provide-style
  [cell workbook]
  (let [style (.getCellStyle cell)]
    (if (nil? style)
      (.createCellStyle workbook)
      style)))

(defn cell-border
  [cell]
  (let [workbook (.. cell getSheet getWorkbook)]
    (.setBorderBottom (provide-style cell workbook) (border :thin))))

(defn row-border
  [row]
  (doall (map cell-border (cell-seq row))))

(defn sep-day
  [rows]
  (when-not (empty? rows)
    (when (different-day? (first rows) (second rows))
      (row-border (first rows)))
    (sep-day (rest rows))))

(defn separate-days
  [sheet]
  (sep-day (row-seq sheet)))

(defn put-sum
  [cell col row-num]
  (.setCellFormula cell (format "SUM(%1$s3:%1$s%2$d)" col row-num)))

(defn add-sums
  [sheet]
  (let [row-num (inc (.getLastRowNum sheet))
        row (.createRow sheet row-num)]
    (put-sum (.createCell row 3) "D" row-num)
    (put-sum (.createCell row 7) "H" row-num)))

(defn save-month
  [data]
  (let [wb (create-workbook "Zeiterfassung" data)
        sheet (select-sheet "Zeiterfassung" wb)
        header-rows (take 2 (row-seq sheet))]
    (do
      (format-cols sheet)
      (separate-days sheet)
      (format-col-size sheet)
      (format-header wb header-rows)
      (add-sums sheet)
      (save-workbook! "2015-05.xlsx" wb))))

(defn -main
  "I don't do a whole lot ... yet."
  [& args]
  
  (println (load-month 5) "Done!"))

