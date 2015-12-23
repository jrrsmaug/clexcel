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

(def here (t/time-zone-for-id "Europe/Vienna"))

(defn month-of
  [entry]
  (t/month
   (tc/to-local-date
    (t/to-time-zone (tc/to-date-time (:datum entry)) here))))

(defn load-month
  [m]
  (filter (fn [entry] (= (month-of entry) m))
          (rest (load-diary))))

(defn map-values-to-vec
  [e]
  (vector
   (:datum e) (:von e) (:bis e) (:dauer e) (:projekt e) (:task e) (:beschreibung e) (:ueberstunden e)))

(defn prepare-for-excel
  "takes a vec of maps and prepares it for an excel sheet"
  [raw]
  (cons ["Datum" "Von" "Bis" "Dauer" "Projekt" "Task" "Beschreibung" "Überstunden"]
        (map map-values-to-vec raw)))

(defn save-month
  [data]
  (let [wb (create-workbook "Zeiterfassung" data)
        sheet (select-sheet "Zeiterfassung" wb)
        header-row (first (row-seq sheet))]
    (do
      (set-row-style! header-row (create-cell-style! wb {:font {:bold true}}))
      (save-workbook! "2015-05.xlsx" wb))))

(defn -main
  "I don't do a whole lot ... yet."
  [& args]
  
  (println (load-month 5) "Done!"))

