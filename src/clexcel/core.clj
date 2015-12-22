(ns clexcel.core
  (:require [clj-time.core :as t])
  (:gen-class))

(use 'dk.ative.docjure.spreadsheet)

(defn load-diary
  "Load a diary into a map"
  []
  (->> (load-workbook "Diary.xls")
       (select-sheet "2015")
       (select-columns {:B :datum :C :von :D :bis :E :dauer :F :projekt :G :task})))

(def here (t/time-zone-for-id "Europe/Vienna"))

(defn month-of
  [entry]
  (t/month
   (clj-time.coerce/to-local-date
    (t/to-time-zone (clj-time.coerce/to-date-time (:datum entry)) here))))

(defn load-month
  [m]
  (filter (fn [entry] (= (month-of entry) m))
          (rest (load-diary))))

(defn save-month
  [data]
  (let [wb (create-workbook "Zeiterfassung"
                            [[data]])
        sheet (select-sheet "Zeiterfassung" wb)
        header-row (first (row-seq sheet))]
    (do
      (save-workbook! "mai.xlsx" wb))))

(defn -main
  "I don't do a whole lot ... yet."
  [& args]
  
  (println (load-month 5) "Done!"))

