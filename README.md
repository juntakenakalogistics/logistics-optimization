# logistics-optimization
Prescriptive logistics optimization via AI (Genetic Algorithm + DBSCAN + DQN) — −57% freight cost on real automotive supply chain data
Prescriptive Logistics Optimization System
AI-driven freight cost and CO₂ reduction for automotive supply chains
![Python](https://img.shields.io/badge/Python-3.11-blue?logo=python&logoColor=white)
![Streamlit](https://img.shields.io/badge/Streamlit-Dashboard-red?logo=streamlit&logoColor=white)
![License](https://img.shields.io/badge/License-MIT-green)
![USP ESALQ](https://img.shields.io/badge/MBA-Data%20Science%20%26%20Analytics-navy)
> **MBA Thesis Project — Universidade de São Paulo (USP/ESALQ), 2023–2026**  
> *"Sistema Prescritivo de Otimização Logística e Inventário de CO₂"*
---
Overview
This system automates and optimizes monthly logistics planning for an automotive parts supply chain, covering ~120 shipments/month from the Warehouse to 71 Tier-1 automotive suppliers across Brazil.
Built entirely in Python (6,626 lines · 117 functions), it replaces manual route and loading planning with a prescriptive analytics pipeline that integrates four optimization algorithms — from classical heuristics to deep reinforcement learning.
Validated on real operational data (Toyota Tsusho Brasil, 2024–2026).
---
Results — Real Operational Data
Metric	Manual Baseline	Best Result (DBSCAN+AG)	Delta
Freight cost	R$ baseline	−57.0%	−57.0%
Total distance	km baseline	−38.5%	−38.5%
CO₂ emissions	ton baseline	−26.7%	−26.7%
Load efficiency	36.5%	86.2%	+49.7 p.p.
Number of trips	14 trips	3 trips	−79%
> *Validated on real delivery data: 21 destinations, 10/05/2026 operational dataset.*
Business impact at scale (~120 trips/month):
Annual freight savings: ~R$ 1.05M
CO₂ reduction: ~110 tons/year
Global Top 5 — Toyota Tsusho Worldwide Kaizen Competition (170+ companies, 50+ countries, Tokyo 2024)
---
System Architecture
```
Input Layer
├── NF-e XML parsing        → automatic extraction of supplier, volume, weight, dimensions
├── Google Maps API         → real road distances and travel times
└── 3D Bin Packing (FFD3D)  → optimal cargo loading per truck

Optimization Layer
├── Algorithm 1: Genetic Algorithm (AG)
├── Algorithm 2: DBSCAN geographic clustering
├── Algorithm 3: DBSCAN + AG hybrid          ← best result
└── Algorithm 4: Deep Q-Network (DQN)        ← reinforcement learning

Output Layer
├── Optimized route plan per algorithm
├── Comparative dashboard (Streamlit)
├── CO₂ inventory per trip
└── Load efficiency and cost breakdown
```
---
Four Optimization Algorithms
1. Genetic Algorithm (AG)
Classical evolutionary optimization applied to the Vehicle Routing Problem (VRP). Encodes routes as chromosomes; uses tournament selection, ordered crossover (OX), and mutation to evolve solutions across generations.
2. DBSCAN Geographic Clustering
Density-Based Spatial Clustering of Applications with Noise. Groups geographically proximate suppliers into delivery clusters before route optimization — reducing inter-cluster travel and enabling parallel loading queues.
3. DBSCAN + AG Hybrid (best result: −57% freight cost)
Two-stage pipeline: DBSCAN clusters destinations by geographic density → AG optimizes the route sequence within and between clusters. Combines spatial intelligence with evolutionary search. Consistently outperforms either algorithm alone.
4. Deep Q-Network (DQN) — Reinforcement Learning
The agent learns a logistics dispatch policy through repeated interaction with a simulated environment. State vector: 12 features (remaining destinations, current load, distance to depot, cluster membership, etc.). Reward function: minimizes total freight cost per episode. Training history is persisted across sessions via replay buffer serialization — enabling warm-start from DBSCAN+AG solutions.
---
Tech Stack
Layer	Technology
Language	Python 3.11
Optimization	Custom GA · DBSCAN (scikit-learn) · DQN (PyTorch)
Routing	Google Maps Distance Matrix API
3D Bin Packing	Custom FFD3D implementation
Data ingestion	XML parsing (lxml) · NF-e schema
Dashboard	Streamlit · Plotly · Folium
Visualization	Three.js (3D cargo view) · Folium (route maps)
Data	Pandas · NumPy
---
Repository Structure
```
logistics-optimization/
│
├── logistica_v90.py          # Main system — 6,626 lines, 117 functions
├── streamlit_app.py          # Web dashboard (6 tabs)
│
├── algorithms/
│   ├── genetic_algorithm.py  # GA implementation
│   ├── dbscan_clustering.py  # DBSCAN geographic clustering
│   ├── hybrid_dbscan_ag.py   # DBSCAN+AG pipeline
│   └── dqn_agent.py          # Deep Q-Network agent + replay buffer
│
├── data/
│   ├── sample_nfe/           # Anonymized NF-e XML examples
│   └── synthetic_dataset.py  # Synthetic data generator for testing
│
├── outputs/
│   ├── route_maps/           # Folium HTML route visualizations
│   ├── load_plans/           # 3D cargo loading diagrams
│   └── results_comparison/   # Algorithm benchmark results
│
├── docs/
│   ├── architecture.png      # System architecture diagram
│   └── dashboard_preview.png # Streamlit dashboard screenshot
│
├── requirements.txt
└── README.md
```
---
Dashboard — 6 Tabs
The Streamlit web application provides real-time comparison across all four algorithms:
Tab	Content
Input	NF-e upload, destination map, cargo summary
Genetic Algorithm	Route plan, load efficiency, cost breakdown
DBSCAN	Cluster map, route visualization
DBSCAN + AG	Hybrid result — best performance
DQN	Training curve, episode rewards, policy result
Comparison	Side-by-side metrics across all algorithms
---
Business Context
Problem: Manual logistics planning for 71 automotive parts suppliers generates suboptimal truck routes, low load efficiency, and excess CO₂ emissions. A human planner typically produces 14 trips for a delivery set that the optimized system resolves in 3.
Scope: ~2,200 m³/month of automotive parts cargo from Port of Santos to Tier-1 suppliers across Brazil.
Regulatory: All NF-e (Nota Fiscal Eletrônica) data handling follows Brazilian tax authority (Receita Federal) schemas. Real operational data is not included in this repository — only anonymized samples and synthetic datasets.
---
Academic Context
This system is the applied thesis (Trabalho de Conclusão de Curso) for the MBA in Data Science & Analytics at Universidade de São Paulo (2023–2026).
Author: Fabio Jun Takenaka
Defense: December 2026
The system was deployed in production at Toyota Tsusho Brasil, Suzano, SP, and validated against 12 months of real operational data.
---
Recognition
🏆 Global Top 5 — Toyota Tsusho Worldwide Kaizen Competition  
Tokyo HQ, 2024 · 170+ companies · 50+ countries  
First South American team to reach the Top 5 in the competition's history.
---
Author
Fabio Jun Takenaka  
Supply Chain Executive · MBA Data Science & Analytics  
25 years · Toyota Group · Brazil & Latin America

License
MIT License — see LICENSE for details.  
Real operational data from Toyota Tsusho Brasil is not included and remains confidential.
