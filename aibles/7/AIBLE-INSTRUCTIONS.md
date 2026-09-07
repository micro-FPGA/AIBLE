# AIBLE Project: Neural Accelerator for Edge AI
**Version:** 1.2
**Date:** 2026-09-07
**Author:** Mistral Medium (LLM)
**Status:** Draft (Invented Example)

---

## 1. Project Overview
**Objective:**
Develop a lightweight, FPGA-based neural accelerator for real-time edge AI inference in low-power devices (e.g., drones, IoT sensors).

**Key Features:**
- Supports INT8 quantization for efficiency.
- Modular design for custom layer integration.
- Compatible with RISC-V soft cores.

---

## 2. Hardware Requirements
- **FPGA Board:** QuickLogic QuickFeather (EOS S3)
- **Toolchain:** Symbiflow + Yosys
- **Peripherals:** UART, SPI, 16MB Flash

---

## 3. Software Dependencies
- **RTOS:** FreeRTOS
- **AI Framework:** TensorFlow Lite for Microcontrollers
- **Compiler:** RISC-V GCC

---

## 4. Design Specifications
### 4.1. Neural Network
- **Architecture:** 3-layer CNN (Input: 64x64 grayscale)
- **Activation:** ReLU
- **Memory:** 512KB on-chip SRAM

### 4.2. FPGA Blocks
- **Accelerator Core:** VexRiscv + Custom MAC Units
- **Clock:** 100MHz (Overclockable to 150MHz)

---

## 5. Workflow Instructions
### Step 1: Clone Repo

git clone https://github.com/micro-FPGA/neural-accel.git
cd neural-accel

### Step 2: Synthesize Design

make SYNTH=1 TARGET=quickfeather
### Step 3: Deploy to FPGA

make FLASH=1 PORT=/dev/ttyUSB0
### Step 4: Test Inference

python3 test_inference.py --model tiny_cnn.tflite --input test_image.png
## 6. Expected Output
Latency: <10ms per inference
Power: <200mW
Accuracy: >90% on MNIST

