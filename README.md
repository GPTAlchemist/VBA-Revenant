# 🧟 VBA Revenant

![Badge: Legacy Code](https://img.shields.io/badge/status-legacy-critical)
![Badge: Built with VBA](https://img.shields.io/badge/tech-VBA-blue)
![Badge: Operational Chaos Slayer](https://img.shields.io/badge/purpose-RFQ%20Intake%20Automation-red)
![Badge: Born 2017](https://img.shields.io/badge/origin-2017-black)

---

## 📚 Table of Contents
- [⚙️ Summary](#️️️️️️️️️️️️️️summary)
- [💣 The Problem](#the-problem)
- [🔨 What I Built](#what-i-built)
- [🧠 What It Proves](#what-it-proves)
- [🚧 The Ugly Truth](#the-ugly-truth)
- [🛠️ What's Next](#whats-next)
- [🔒 Disclaimer](#disclaimer)

---

## ⚙️ Summary

**VBA Revenant** is an undead automation engine I built in 2017 using Excel, the macro recorder, and 600 gallons of coffee. Its job: stop estimators from screwing up RFQ intake. It wasn’t clean. It wasn’t modular. It wasn’t even safe. But it worked — and it saved hours of rework every week in a live estimating environment at a defense contractor.

This was built without formal training, without code review, and without a safety net.

---

## 💣 The Problem

RFQ intake at kSARIA was a mess:
- Estimators missed fields
- Work items were misnamed
- Templates got duplicated, broken, or left out
- Critical dates were skipped

Every proposal became a minefield. I watched it happen repeatedly — and nobody fixed it.

So I did.

---

## 🔨 What I Built

A full Excel-based automation suite inside a single macro. It:
- Prompted the user for every required intake field (POC, dates, company, ship info)
- Dynamically cloned and renamed proposal sheets for each work item
- Inserted live formulas into summary tables and cost rollups
- Toggled visibility on support sheets based on scope
- Handled travel logic with conditional generation
- Refreshed pivot tables and structured the output consistently

And all of that was done inside a single subroutine: `StartUp`. Because I didn’t know any better — and didn’t care. It got the job done.

---

## 🧠 What It Proves

- I don’t need ideal conditions to solve real problems
- I can turn vague workflows into executable logic
- I will finish what I start — even when I have to teach myself from scratch
- I can design tools that fix operational chaos under pressure

This isn’t “learning to code.” This is *applying logic with no support and no margin for failure.*

---

## 🚧 The Ugly Truth

This code is cursed:
- One giant macro — zero modularization
- `.Activate` and `.Select` everywhere
- Magic cell references like `Cells(9,2)` littered across the battlefield
- GOTO loops that would give a senior dev a stroke
- All user interaction done with `InputBox`, over and over and over again

But here’s the thing: it worked — and it saved the department hours of manual labor every week.

---

## 🛠️ What’s Next

This tool’s legacy lives on — but it’s time for a rebuild. I’m planning to:
- Rewrite this in Python using `openpyxl` or `xlwings`
- Build a proper UI (Flask or tkinter) to replace `InputBox` hell
- Modularize the logic: sheet generation, validation, summary injection
- Add YAML config control and maybe even GPT-driven input validation

This project won’t be hidden anymore. It’s part of the foundation for everything I’ve built since — including GPT-integrated systems that now replace ops failure with automation you can trust.

---

## 🔒 Disclaimer

This README was generated using the [README Synth GPT](https://github.com/GPTAlchemist/README-Synth), a tool designed to convert user-authored documentation, design logic, and development notes into clear, publishable Markdown.

All ideas, descriptions, and feature logic originated from the creator of this tool.  
README Synth GPT structured, refined, and formatted the content — but it did not invent the product, its claims, or its language.

For full transparency on how this system works, see the GitHub project: [README Synth GPT →](https://github.com/GPTAlchemist/README-Synth)

---

**Created by Andrew Polk**  
Estimating Manager → AI Workflow Architect  
[GitHub](https://github.com/GPTAlchemist) • [LinkedIn](https://linkedin.com/in/andrew-l-polk)
