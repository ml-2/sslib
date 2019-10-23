(ns sslib.core
  (:require [clojure.spec.alpha :as s])
  (:use dataorigin.core)
  (:import
   [org.apache.poi.ss.usermodel
    WorkbookFactory Workbook Sheet Row Cell CellType DateUtil]
   [org.apache.poi.xssf.usermodel XSSFWorkbook]
   [java.util Locale Date]
   [java.io File FileInputStream FileOutputStream]
   [java.time LocalDateTime LocalDate ZoneOffset]
   [clojure.lang ExceptionInfo])
  (:gen-class))

(derive ::error-cell ::ss-error)
(derive ::formula-cell ::ss-error)
(derive ::unknown-cell-type ::ss-error)
(derive ::fail-mapify-row ::ss-error)
(derive ::fail-mapify-sheet ::ss-error)

(def letters "ABCDEFGHIJKLMNOPQRSTUVWXYZ")

;; :as-string, :error
(def ^:dynamic formula-cell-handling :error)
;; :as-number, :error
(def ^:dynamic error-cell-handling :error)

(defn column-name [cell-num]
  (loop [cell-num (int cell-num)
         name ()]
    (cond
      (zero? cell-num)
      (apply str name)
      (< cell-num 0)
      "_"
      true
      (recur (int (/ cell-num 26))
             (cons (nth letters (rem cell-num 26)) name)))))

(defn row-col-str [row-num cell-num]
  (let [column (if (nil? cell-num)
                 "_"
                 (column-name cell-num))]
    (str column (inc row-num))))

(defn origin-to-error-message [origin]
  (cond
    (nil? origin)
    ""
    (= (:origin-kind origin) ::excel-workbook)
    (str (:sheet-name origin) ":"
         (row-col-str (:row-num origin) (:cell-num origin)))))

;; TODO: ::rest-validators checking, if it's present
(s/def ::columns (fn [columns]
                   (if-not (and (contains? columns ::ordering)
                                (contains? columns ::validators)
                                (coll? (::ordering columns))
                                (map? (::validators columns)))
                     false
                     (reduce #(and %1 (contains? (::validators columns) %2))
                             true (::ordering columns)))))

(defn make-identity-validators [ordering]
  (apply hash-map (interleave ordering (repeat identity))))

(defn make-validator [f ex]
  (fn [data]
    (if (f (unwrap data))
      data
      (throw (ex data)))))

(defn remove-trailing
  ([coll]
   (remove-trailing nil? coll))
  ([fn coll]
   (reverse (drop-while fn (reverse coll)))))

(defn null-row? [row]
  (or (empty? row)
      (every? #(nil? (unwrap %)) row)))

(defmulti make-local-date-time (fn [date] (type date)))
(defmethod make-local-date-time Double [date]
  (make-local-date-time (DateUtil/getJavaDate date)))
(defmethod make-local-date-time Date [date]
  (LocalDateTime/ofInstant (.toInstant date)
                           ZoneOffset/UTC))
(defmethod make-local-date-time LocalDateTime [date]
  date)
(defmethod make-local-date-time LocalDate [date]
  (.atStartOfDay date))

(defn maybe-make-local-date-time [date]
  (try
    (make-local-date-time date)
    (catch IllegalArgumentException ex
      nil)))

(defn load-cell [origin ^Cell cell]
  (let [cell-data
        (if (nil? cell)
          nil
          (let [e (.getCellTypeEnum cell)]
            (cond
              (= e CellType/_NONE)
              nil
              (= e CellType/BLANK)
              nil
              (= e CellType/ERROR)
              ;; TODO: Handle this well: give a proper value
              (cond
                (= error-cell-handling :as-number)
                (.getErrorCellValue cell)
                (or true (= error-cell-handling :error))
                (throw (ex-info "Error Cell" {:data (make-data cell origin)
                                              :kind ::error-cell})))
              (= e CellType/FORMULA)
              (cond
                (= formula-cell-handling :as-string)
                (.getCellFormula cell)
                (or true (= formula-cell-handling :error))
                (throw (ex-info "Formula Cell" {:data (make-data cell origin)
                                                :kind ::formula-cell})))
              (= e CellType/BOOLEAN)
              (.getBooleanCellValue cell)
              (= e CellType/STRING)
              (.getStringCellValue cell)
              (= e CellType/NUMERIC)
              (if (DateUtil/isCellDateFormatted cell)
                (make-local-date-time (.getDateCellValue cell))
                (let [d (.getNumericCellValue cell)]
                  (if (= (mod d 1) 0.0)
                    (long d)
                    d)))
              true
              (throw (ex-info "Unreachable"
                              {:data (make-data nil origin)
                               :kind ::unknown-cell-type})))))]
    (make-data cell-data (assoc origin :origin-data ::cell))))

(defn load-row [origin ^Row row]
  (loop [i 0
         coll (transient [])]
    (if (or (nil? row) (>= i (.getLastCellNum row)))
      (make-data (persistent! coll) (assoc origin :origin-data ::row))
      (recur (inc i)
             (conj! coll (load-cell (assoc origin :cell-num i)
                                    (.getCell row i)))))))

(defn load-sheet [origin ^Sheet sheet]
  (loop [i 0
         coll (transient [])]
    (if (> i (.getLastRowNum sheet))
      (make-data (persistent! coll) (assoc origin :origin-data ::sheet))
      (recur (inc i)
             (conj! coll (load-row (assoc origin :row-num i)
                                   (.getRow sheet i)))))))

(defn load-workbook [origin ^Workbook workbook]
  (loop [sheets []]
    (if (>= (count sheets) (.getNumberOfSheets workbook))
      (make-data sheets (assoc origin :origin-data ::workbook))
      (let [sheet (.getSheetAt workbook (count sheets))
            sheet-name (.getSheetName sheet)]
        (recur
         (conj sheets
               (load-sheet
                (assoc origin :sheet-name sheet-name) sheet)))))))

;; TODO: Way to force loading only select sheets, to handle bad sheets mixed with
;; good sheets better.
(defn load-workbook! [file-name]
  (let [file (File. file-name)
        stream (FileInputStream. file)
        workbook (WorkbookFactory/create stream)
        origin {:origin-kind ::excel-workbook
                :file-name file-name}
        ret (load-workbook origin workbook)]
    (.close workbook)
    ret))

;; TODO: This was sort of copied from docjure (MIT license)
(defmulti set-cell! (fn [^Cell cell data] (type data)))

(defmethod set-cell! String [^Cell cell data]
  (.setCellValue cell ^String data))

(defmethod set-cell! Number [^Cell cell data]
  (.setCellValue cell ^Double (double data)))

(defmethod set-cell! Boolean [^Cell cell data]
  (.setCellValue cell ^Boolean data))

(defmethod set-cell! LocalDateTime [^Cell cell data]
  (.setCellValue cell ^Date (Date/from (.toInstant data ZoneOffset/UTC))))
;; TODO: The following is a pretty crude timezone hack:
(defmethod set-cell! LocalDate [^Cell cell data]
  (.setCellValue cell ^Date (Date/from (.toInstant
                                        (.atTime data 12 0)
                                        ZoneOffset/UTC))))
(defmethod set-cell! nil [^Cell cell data]
  (let [^String null nil]
    (.setCellValue cell null)))

(defn save-row! [row values]
  (loop [vals (seq values)
         i 0]
    (if-not (empty? vals)
      (do
        (set-cell! (.createCell row i) (unwrap (first vals)))
        (recur (rest vals)
               (inc i))))))


(defn save-cell! [cell data]
  (set-cell! cell (unwrap data)))

(defn save-map-row! [row map-row columns]
  {:pre [(s/valid? ::columns columns)]}
  (let [cells (unwrap map-row)
        validators (::validators columns)]
    (loop [i 0 ordering (::ordering columns)]
      (if (empty? ordering)
        (if-not (::rest-validators columns)
          nil
          (let [cell-validator (::cell-rest-validator
                                (::rest-validators columns))
                all-validator (::all-rest-validator
                               (::rest-validators columns))]
          (loop [i i
                 rst (all-validator (or (::rest cells) ()))]
            (if (empty? rst)
              nil
              (do
                (save-cell! (.createCell row i) (cell-validator (first rst)))
                (recur (inc i) (rest rst)))))))
        (do
          (save-cell! (.createCell row i) ((get validators (first ordering))
                                           (get cells (first ordering))))
          (recur (inc i) (rest ordering)))))))

(defn save-map-sheet!
  ([sheet map-sheet columns] (save-map-sheet! sheet map-sheet columns nil))
  ([sheet map-sheet columns titles]
   (let [i 0
         i
         (if titles
           (let [row (.createRow sheet 0)]
             (save-row! row (unwrap titles))
             (inc i))
           i)
         rows (unwrap map-sheet)]
     (loop [i i
            rows-rest rows]
       (if (empty? rows-rest)
         nil
         (let [row (.createRow sheet i)]
           (save-map-row! row (first rows-rest) columns)
           (recur (inc i) (rest rows-rest))))))))

(defn unwrap-map-row [row columns]
  (loop [columns (::ordering columns)
         row (unwrap row)]
    (if (empty? columns)
      (if (::rest row)
        (update row ::rest #(map unwrap %))
        row)
      (recur (rest columns)
             (update row (first columns) unwrap)))))

(defn unwrap-sheet [sheet]
  (map #(map unwrap (unwrap %)) (unwrap sheet)))

(defn unwrap-map-sheet [sheet columns]
  (map #(unwrap-map-row % columns) (unwrap sheet)))

;; TODO: Possibly create the map, then validate the stuff.
;; This would make validation when saving easier, too. And make validation
;; easier to disable.
;; TODO: Option to skip title row
(defn mapify
  ([sheet columns] (mapify sheet columns nil))
  ([sheet columns keep-titles]
   {:pre [(s/valid? ::columns columns)]}
   ;; TODO: sheet and :post
   (try
     (make-data (doall
                 (map
                  (fn [row]
                    (let [urow (remove-trailing null-row? (unwrap row))]
                      (try
                        (loop [i 0
                               hmap {}]
                          (if (>= i (count (::ordering columns)))
                            (let [cell-validator (::cell-rest-validator
                                                  (::rest-validators columns))
                                  all-validator (::all-rest-validator
                                                 (::rest-validators columns))]
                              (make-data (if (::rest-validators columns)
                                           (assoc hmap ::rest (all-validator
                                                               (doall
                                                                (map
                                                                 cell-validator
                                                                 (remove-trailing
                                                                  nil? (drop i urow))))))
                                           hmap)
                                         (origin row)))
                            (recur
                             (inc i)
                             (let [column (nth (::ordering columns) i)
                                   validate (get (::validators columns) column)]
                               (assoc hmap
                                      column
                                      (validate (if (< i (count urow))
                                                  (nth urow i)
                                                  (load-cell
                                                   (assoc (origin row) :cell-num i)
                                                   nil))))))))
                        (catch ExceptionInfo ex
                          (throw (ex-info "Failed to mapify row"
                                          {:data (make-data nil (origin row)) :kind ::fail-mapify-row}
                                          ex))))))
                  (remove (fn [row]
                            (loop [skip-checkers (::skip-checkers columns)]
                              (cond
                                (empty? skip-checkers)
                                false
                                ((first skip-checkers) (unwrap row))
                                true
                                true
                                (recur (rest skip-checkers)))))
                          (remove-trailing
                           #(null-row? (unwrap %))
                           (if keep-titles
                            (unwrap sheet)
                            (rest (unwrap sheet))))))) (origin sheet))
     (catch ExceptionInfo ex
       (throw (ex-info "Failed to mapify sheet"
                       {:data (make-data nil (origin sheet)) :kind ::fail-mapify-sheet}
                       ex))))))


;; An extractor is a function that takes a workbook and a state, where the state
;; is {:position, :data, :message}. The extractor returns a new state.

;; :position is a {:sheet, :row, :column}. :data is a map of extracted data.
;; :message is a message for the caller.

;; An extractor can call another extractor internally. This is how loops are
;; implemented.

;; Extractors work on sheets that aren't mapified.

(defn get-workbook-position [workbook position]
  (let [sheets (unwrap workbook)]
    (if (>= (:sheet position) (count sheets))
      nil
      (let [rows (unwrap (nth sheets (:sheet position)))]
        (if (>= (:row position) (count rows))
          nil
          (let [columns (unwrap (nth rows (:row position)))]
            (if (>= (:column position) (count columns))
              nil
              (nth columns (:column position)))))))))

(defn update-extractor-state [state, position-merge, message, data-updater]
  (update (assoc state
                 :position (merge (:position state) position-merge)
                 :message message)
          :data data-updater))

(defn apply-extractors
  ([extractors]
   #(apply-extractors extractors %1 %2))
  ([extractors workbook, state]
   (reduce (fn [state fun]
             (let [ret-state (fun workbook state)]
               (if (not= (:message ret-state) :found)
                 (throw (ex-info "Extractor returned a message other than :found"
                                 {:state ret-state :extractor fun}))
                 ret-state)))
           state extractors)))
