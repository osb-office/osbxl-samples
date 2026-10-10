;; WebAssembly Text Format (WAT) sample for osbxl calculation engine integration
;; Demonstrates fast in-memory cell calculation bridge for wasmosbxl runtime
(module
  ;; Memory declaration for spreadsheet calculation buffer
  (memory (export "memory") 1)

  ;; Exported calculation helper function: computeCellTotal
  ;; Takes unitPrice (i32) and quantity (i32), returns total (i32)
  (func $computeCellTotal (param $unitPrice i32) (param $quantity i32) (result i32)
    local.get $unitPrice
    local.get $quantity
    i32.mul
  )
  (export "computeCellTotal" (func $computeCellTotal))

  ;; Exported calculation helper function: computeTax
  ;; Computes tax amount with 10% rate
  (func $computeTax (param $subtotal i32) (result i32)
    local.get $subtotal
    i32.const 10
    i32.mul
    i32.const 100
    i32.div_s
  )
  (export "computeTax" (func $computeTax))
)
