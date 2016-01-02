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

(defn month-of
  "Returns the month of the datum cell of a row"
  [entry]
  (t/month (tc/to-local-date (timezoned (:datum entry)))))

(defn timezoned
  [date]
  (when-not (nil? date) (t/to-time-zone (tc/to-date-time date) tz-here)))

(defn load-month
  [m]
  (filter (fn [entry] (= (month-of entry) m))
          (rest (load-diary))))

(defn map-values-to-vec
  [e]
  (row-vec [:datum :von :bis :dauer :projekt :task :beschreibung :ueberstunden] e))

(defn prepare-for-excel
  "takes a vec of maps and prepares it for an excel sheet"
  [raw]
  (cons ["Datum" "Von" "Bis" "Dauer" "Projekt" "Task" "Beschreibung" "Überstunden"]
        (map map-values-to-vec (fix-times raw))))

(defn fix-times
  [raw]
  (map fix-time-bis (map fix-time-von raw)))

(defn fix-time-von
  [entry]
  (update-in entry [:von] fix-t))

(defn fix-time-bis
  [entry]
  (update-in entry [:bis] fix-t))

(defn fix-t
  [time]
  (tc/to-date (t/plus (timezoned time) (t/days 1))))

(defn save-month
  [data]
  (let [wb (create-workbook "Zeiterfassung" data)
        sheet (select-sheet "Zeiterfassung" wb)
        header-row (first (row-seq sheet))]
    (do
      (set-row-style! header-row (create-cell-style! wb {:font {:bold true}}))
      (save-workbook! "2015-05.xlsx" wb))))

(defn col-seq
  "Returns a sequence of all cells in column col of the sheet"
  [col sheet]
  (seq (map #(.getCell % col) (row-seq sheet))))

(defn -main
  "I don't do a whole lot ... yet."
  [& args]
  
  (println (load-month 5) "Done!"))

