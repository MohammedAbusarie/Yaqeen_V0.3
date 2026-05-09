Production AI Engineering Rules



These rules govern how systems must be designed, extended, and maintained in production environments.  

They apply to any project regardless of language, framework, or domain.



Violations are considered architectural risks and must be corrected before proceeding.



\--------------------------------------------------

0\. Project Initialization and Architectural Responsibility

\--------------------------------------------------



DON’T begin implementation without understanding the system structure.



DON’T assume an architecture exists.



If the project is new and no architecture or system design is defined:



&#x20;   You must establish a well-engineered architecture before writing implementation logic.

&#x20;   The architecture must be scalable, maintainable, testable, and observable.

&#x20;   The architecture must define clear boundaries, responsibilities, and dependency flows.



If an architecture already exists:



&#x20;   DON’T introduce patterns, structures, or flows that conflict with the current design.

&#x20;   DON’T refactor core architecture without explicit justification.

&#x20;   DON’T bypass established layers, contracts, or abstractions.

&#x20;   All enhancements, fixes, and features must preserve architectural integrity.



Architecture consistency is more important than implementation speed.



\--------------------------------------------------

1\. Documentation and Communication Standards

\--------------------------------------------------



DON’T use inline comments to explain behavior.



DON’T rely on scattered comments to describe system logic.



Instead:



&#x20;   Use clear, structured, well-written docstrings to describe:

&#x20;       - Purpose

&#x20;       - Inputs

&#x20;       - Outputs

&#x20;       - Side effects

&#x20;       - Failure conditions

&#x20;       - Dependencies



DON’T leave undocumented behavior in production systems.



DON’T assume future maintainers will infer intent.



System behavior must be self-explanatory through structured documentation.



\--------------------------------------------------

2\. Reliability and Operational Readiness

\--------------------------------------------------



DON’T implement functionality without visibility into its behavior.



DON’T deploy logic that cannot be observed, traced, or diagnosed.



DON’T allow failures to occur silently.



DON’T assume operations will always succeed.



DON’T omit structured logging.



DON’T omit error handling.



DON’T omit retry or recovery mechanisms where failure is possible.



DON’T ignore system health signals.



Every critical operation must be:



&#x20;   Observable  

&#x20;   Measurable  

&#x20;   Recoverable  

&#x20;   Diagnosable  



\--------------------------------------------------

3\. Security and Access Control

\--------------------------------------------------



DON’T expose system capabilities without access control.



DON’T trust external input.



DON’T assume internal traffic is safe.



DON’T allow unbounded usage of system resources.



DON’T allow sensitive operations to occur without traceability.



Security must be enforced consistently across all entry points.



\--------------------------------------------------

4\. Architectural Consistency

\--------------------------------------------------



DON’T mix architectural styles within the same module or workflow.



DON’T introduce new patterns without evaluating compatibility.



DON’T violate separation of concerns.



DON’T bypass defined layers.



DON’T access infrastructure directly when abstractions exist.



DON’T create hidden dependencies.



DON’T allow inconsistent naming conventions.



DON’T change conventions mid-project.



Consistency is a system property, not a coding preference.



\--------------------------------------------------

5\. Distributed System and Environment Awareness

\--------------------------------------------------



DON’T assume local behavior represents production behavior.



DON’T hardcode environment-specific values.



DON’T ignore network variability.



DON’T assume infinite availability.



DON’T assume zero latency.



DON’T assume external services are reliable.



Systems must tolerate:



&#x20;   Delays  

&#x20;   Failures  

&#x20;   Partial availability  

&#x20;   Load spikes  



\--------------------------------------------------

6\. Scalability and Performance Discipline

\--------------------------------------------------



DON’T design logic that scales linearly with data size without evaluation.



DON’T perform repetitive resource operations unnecessarily.



DON’T create unbounded resource consumption.



DON’T ignore performance costs.



DON’T assume small workloads remain small.



DON’T allow inefficient access patterns to persist.



DON’T deploy systems that degrade unpredictably under load.



Performance must be intentional, not accidental.



\--------------------------------------------------

7\. Resource and Cost Responsibility

\--------------------------------------------------



DON’T create operations with uncontrolled execution frequency.



DON’T repeat expensive operations without necessity.



DON’T ignore infrastructure cost impact.



DON’T assume compute resources are unlimited.



DON’T allow resource usage to grow without visibility.



Every system must operate within predictable resource boundaries.



\--------------------------------------------------

8\. Context and Scope Discipline

\--------------------------------------------------



DON’T load unnecessary information into working context.



DON’T duplicate information already available.



DON’T expand scope beyond the task requirement.



DON’T generate excessive output without need.



DON’T introduce complexity without measurable benefit.



Systems must remain focused and minimal.



\--------------------------------------------------

9\. Design Before Implementation

\--------------------------------------------------



DON’T write implementation logic before defining structure.



DON’T build behavior before defining responsibilities.



DON’T create components without defined interfaces or contracts.



DON’T introduce dependencies without documenting relationships.



DON’T allow implicit system behavior.



Every component must have:



&#x20;   A defined responsibility  

&#x20;   A defined interface  

&#x20;   A defined dependency boundary  



Design precedes implementation.



Always.



\--------------------------------------------------

10\. Change and Extension Discipline

\--------------------------------------------------



DON’T modify system behavior without understanding its impact.



DON’T introduce changes that break existing flows.



DON’T alter system boundaries without justification.



DON’T implement features that bypass validation or consistency checks.



DON’T prioritize speed over stability.



System evolution must be controlled, predictable, and reversible.



\--------------------------------------------------

11\. Failure Management

\--------------------------------------------------



DON’T assume success.



DON’T ignore failure paths.



DON’T allow cascading failures.



DON’T allow a single failure to collapse the system.



DON’T leave the system in an undefined state after failure.



Systems must fail safely.



\--------------------------------------------------

12\. Maintainability Requirements

\--------------------------------------------------



DON’T create logic that only its author can understand.



DON’T introduce tightly coupled components.



DON’T hide behavior in unexpected locations.



DON’T create structures that resist modification.



DON’T sacrifice clarity for cleverness.



Maintainability is a primary system requirement.



\--------------------------------------------------

13\. Enforcement Principle

\--------------------------------------------------



DON’T generate implementation that violates system rules.



DON’T proceed when architecture is unclear.



DON’T continue when system integrity is at risk.



If a rule conflict is detected:



&#x20;   Stop  

&#x20;   Identify the conflict  

&#x20;   Resolve the design  

&#x20;   Then proceed  



System correctness takes priority over progress.



\--------------------------------------------------

Core Principle



Build systems that survive change.

Not systems that only work today.

